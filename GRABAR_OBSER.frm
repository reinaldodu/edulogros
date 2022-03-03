VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GRABAR_OBSER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grabar Observaciones"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9255
   Icon            =   "GRABAR_OBSER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   6600
      Picture         =   "GRABAR_OBSER.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprime la información de la lista"
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "DEL"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      ToolTipText     =   "Borrar un estudiante de la lista"
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "INS"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      ToolTipText     =   "Actualizar la lista de estudiantes"
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox Text7 
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
      Height          =   360
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      Picture         =   "GRABAR_OBSER.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guarda la información de la lista"
      Top             =   5880
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
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
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "GRABAR_OBSER.frx":0EA6
      Left            =   7680
      List            =   "GRABAR_OBSER.frx":0EB9
      TabIndex        =   0
      Text            =   "PRIMERO"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   5640
      Width           =   5775
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   320
         Left            =   5280
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Materia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   9015
      Begin VB.CommandButton Command8 
         Caption         =   "Pegar columna"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Ver observaciones"
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Pegar &todo"
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "P&egar fila"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Copiar"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4680
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid MATI12 
         Height          =   4215
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7435
         _Version        =   393216
         Rows            =   1
         Cols            =   14
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total estudiantes..."
         Height          =   195
         Left            =   6960
         TabIndex        =   19
         Top             =   4680
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opciones de Lista"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   7200
      TabIndex        =   17
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6360
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5880
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6840
      TabIndex        =   14
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "GRABAR_OBSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim simucopy(10) As String, kcpy As Integer, vcpy As Integer, VerInfo As Boolean, CuentaObs As Integer

'Función que devuelve el # del registro de la observación dada
Public Function RegObserv(NumObser As Integer) As Integer
Dim Contador As Integer
Contador = 0
RegObserv = 0
NAR = FreeFile
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    RegObserv = RegObserv + 1
    Get #NAR, RegObserv, logru
    If Trim(logru.indicador) <> "L" Then
        Contador = Contador + 1
    End If
    If Contador = NumObser Then
        Close #NAR
        'NAR = NAR - 1
        Exit Function
    End If
Wend
Close #NAR
'NAR = NAR - 1
End Function

'Función que devuelve el #(orden) de la observacion según el registro dado
Public Function OrdenObserv(NumReg As Integer) As Integer
Dim Contador As Integer
Contador = 0
OrdenObserv = 0
NAR = FreeFile
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    Contador = Contador + 1
    Get #NAR, Contador, logru
    If Trim(logru.indicador) <> "L" Then
        OrdenObserv = OrdenObserv + 1
    End If
    If Contador = NumReg Then
        Close #NAR
        'NAR = NAR - 1
        Exit Function
    End If
Wend
Close #NAR
'NAR = NAR - 1
End Function

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

If VALI4 = False Then
    Call Command6_Click
End If
Unload Ver_Obser
MATI12.ToolTipText = ""
MATI12.Rows = 1
Label4.Caption = ""
Label6.Caption = ""
Frame1.Caption = ""
Frame2.Caption = ""
Label1.Caption = ""
Text7.Text = 0
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command2.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
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
    If RTrim(argra.grado) = RE22 And RTrim(argra.nom_grup) = Combo2.Text And argra.num_area = que Then
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
If RTrim(Combo1.Text) = "FINAL" Then
    lw = 5
End If


If JOJI = "UNICA" Then
fl = "1"
End If
If JOJI = "MAÑANA" Then
fl = "2"
End If
If JOJI = "TARDE" Then
fl = "3"
End If
If JOJI = "NOCHE" Then
fl = "4"
End If
ser = Left(RE22, 3)
FERT = 0
CuentaObs = 0
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    FERT = FERT + 1
    Get #NAR, FERT, logru
    If Trim(logru.indicador) <> "L" Then
        CuentaObs = CuentaObs + 1
    End If
Wend
Close #NAR


Command3.Enabled = True
Command2.Enabled = True
' Se verifica si está bloqueado el periodo para no habilitar el botón guardar
'If VeriPeriodo(lw) = False Then
'    Command6.Enabled = False
'     MsgBox "EL PERIODO " & Combo1 & " SOLO ESTA DISPONIBLE PARA CONSULTA", 32, "Grabar observaciones"
'Else
    Command6.Enabled = True
'End If
Command7.Enabled = True
Command9.Enabled = True
Command10.Enabled = True

Label4.Caption = Combo2.Text
Label6.Caption = Combo2.Text & que & lw
If Dir(Ruta & Combo2.Text & que & lw & ".obs") = "" Then
    For I = 1 To (ret - 1)
        MATI12.Rows = I + 1
        MATI12.TextMatrix(I, 0) = I
        Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
        Get #NAR, I, alugru
        Close #NAR
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        MATI12.TextMatrix(I, 12) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        MATI12.TextMatrix(I, 13) = alumno.n_carnet
    Next I
Else
    Y = 0
    Open Ruta & Combo2.Text & que & lw & ".obs" For Random As #NAR Len = Len(notas)
    While Not EOF(NAR)
        Y = Y + 1
        Get #NAR, Y, notas
    Wend
    Close #NAR
    For I = 1 To (Y - 1)
        MATI12.Rows = I + 1
        MATI12.TextMatrix(I, 0) = I
        Open Ruta & Combo2.Text & que & lw & ".obs" For Random As #NAR Len = Len(notas)
        Get #NAR, I, notas
        Close #NAR
        For J = 1 To 10
            If notas.area(J) = 0 Then
                MATI12.TextMatrix(I, J) = ""
            Else
                MATI12.TextMatrix(I, J) = OrdenObserv(notas.area(J))
            End If
        Next J
        'MATI12.TextMatrix(I, 11) = RTrim(notas.JV)
        If notas.FA = 0 Then
            MATI12.TextMatrix(I, 11) = ""
        Else
            MATI12.TextMatrix(I, 11) = notas.FA
        End If
        If RTrim(notas.num_carnet) = "" Then
            GoTo salbla
        End If
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(notas.num_carnet)), alumno
        Close #NAR
        MATI12.TextMatrix(I, 12) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        MATI12.TextMatrix(I, 13) = alumno.n_carnet
salbla:
    Next I
End If

Label1.Caption = "GRABACION DE OBSERVACIONES JORNADA:" & JOJI & "  GRADO: " & RE22
Frame1.Caption = "GRUPO: " & Combo2.Text & " - " & " AREA: " & Combo3.Text & " - " & " PROFESOR(A): " & PRO
Frame2.Caption = "PERIODO " & Combo1.Text
Text7.Text = I - 1
MATI12.Row = 1
MATI12.Col = 1
MATI12.SetFocus
Screen.MousePointer = 0


'If VALI4 = False Then
'    Call Command6_Click
'End If
'Unload Ver_Obser
'MATI12.ToolTipText = ""
'MATI12.Rows = 1
'Label4.Caption = ""
'Label6.Caption = ""
'Frame1.Caption = ""
'Frame2.Caption = ""
'Label1.Caption = ""
'Text7.Text = 0
'Command3.Enabled = False
'Command4.Enabled = False
'Command5.Enabled = False
'Command2.Enabled = False
'Command6.Enabled = False
'Command7.Enabled = False
'Command9.Enabled = False
'Command10.Enabled = False
'If Dir(Ruta & Combo2.Text & ".gru") = "" Then
'    MsgBox "GRUPO INCORRECTO", 48
'    Exit Sub
'End If
'Screen.MousePointer = 11
'NAR = FreeFile
'TN = 0
'Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
'While Not EOF(NAR)
'    TN = TN + 1
'    Get #NAR, TN, mate
'    If RTrim(mate.nom) = Combo3.Text Then
'        que = mate.num
'    End If
'Wend
'Close #NAR
'ret = 0
'Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
'While Not EOF(NAR)
'    ret = ret + 1
'    Get #NAR, ret, alugru
'Wend
'Close #NAR
'NAR = FreeFile
'Open Ruta & "infcur.edu" For Input As #NAR
'While Not EOF(NAR)
'    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
'    If RTrim(icur.nom) = RTrim(Combo2.Text) Then
'        RE22 = RTrim(icur.grado)
'        JOJI = RTrim(icur.jornada)
'        GoTo ALTU33
'    End If
'Wend
'ALTU33:
'Close #NAR
'pio = 0
'cona = 0
'Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
'While Not EOF(NAR)
'    cona = cona + 1
'    Get #NAR, cona, argra
'    If RTrim(argra.grado) = RE22 And RTrim(argra.nom_grup) = Combo2.Text And argra.num_area = que Then
'        NAR = FreeFile
'        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
'        Get #NAR, (argra.num_pro), profe
'        Close #NAR
'        PRO = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
'        pio = 1
'        NAR = NAR - 1
'    End If
'Wend
'Close #NAR
'If pio = 0 Then
'MsgBox "NO SE HA CREADO EL AREA " & Combo3.Text & " PARA ESTE GRUPO", 64, "ADVERTENCIA"
'    Combo3.SetFocus
'    Screen.MousePointer = 0
'    Exit Sub
'End If
'Command3.Enabled = True
'Command2.Enabled = True
'Command6.Enabled = True
'Command7.Enabled = True
'Command9.Enabled = True
'Command10.Enabled = True
'If RTrim(Combo1.Text) = "PRIMERO" Then
'    lw = 1
'End If
'If RTrim(Combo1.Text) = "SEGUNDO" Then
'    lw = 2
'End If
'If RTrim(Combo1.Text) = "TERCERO" Then
'    lw = 3
'End If
'If RTrim(Combo1.Text) = "CUARTO" Then
'    lw = 4
'End If
'If RTrim(Combo1.Text) = "FINAL" Then
'    lw = 5
'End If
'Label4.Caption = Combo2.Text
'Label6.Caption = Combo2.Text & que & lw
'If Dir(Ruta & Combo2.Text & que & lw & ".obs") = "" Then
'    For I = 1 To (ret - 1)
'        MATI12.Rows = I + 1
'        MATI12.TextMatrix(I, 0) = I
'        Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
'        Get #NAR, I, alugru
'        Close #NAR
'        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'        Get #NAR, (Val(alugru.num_carnet)), alumno
'        Close #NAR
'        MATI12.TextMatrix(I, 12) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
'        MATI12.TextMatrix(I, 13) = alumno.n_carnet
'    Next I
'Else
'    Y = 0
'    Open Ruta & Combo2.Text & que & lw & ".obs" For Random As #NAR Len = Len(notas)
'    While Not EOF(NAR)
'        Y = Y + 1
'        Get #NAR, Y, notas
'    Wend
'    Close #NAR
'    For I = 1 To (Y - 1)
'        MATI12.Rows = I + 1
'        MATI12.TextMatrix(I, 0) = I
'        Open Ruta & Combo2.Text & que & lw & ".obs" For Random As #NAR Len = Len(notas)
'        Get #NAR, I, notas
'        Close #NAR
'        For J = 1 To 10
'            If notas.area(J) = 0 Then
'                MATI12.TextMatrix(I, J) = ""
'            Else
'                MATI12.TextMatrix(I, J) = notas.area(J)
'            End If
'        Next J
'        'MATI12.TextMatrix(I, 11) = RTrim(notas.FA)
'        MATI12.TextMatrix(I, 11) = notas.FA
'        If RTrim(notas.num_carnet) = "" Then
'            GoTo salbla
'        End If
'        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'        Get #NAR, (Val(notas.num_carnet)), alumno
'        Close #NAR
'        MATI12.TextMatrix(I, 12) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
'        MATI12.TextMatrix(I, 13) = alumno.n_carnet
'salbla:
'    Next I
'End If
'If JOJI = "UNICA" Then
'fl = "1"
'End If
'If JOJI = "MAÑANA" Then
'fl = "2"
'End If
'If JOJI = "TARDE" Then
'fl = "3"
'End If
'If JOJI = "NOCHE" Then
'fl = "4"
'End If
'ser = Left(RE22, 3)
'FERT = 0
'Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
'While Not EOF(NAR)
'    FERT = FERT + 1
'    Get #NAR, FERT, logru
'Wend
'Close #NAR
'Label1.Caption = "GRABACION DE OBSERVACIONES JORNADA:" & JOJI & "  GRADO: " & RE22
'Frame1.Caption = "GRUPO: " & Combo2.Text & " - " & " AREA: " & Combo3.Text & " - " & " PROFESOR(A): " & PRO
'Frame2.Caption = "PERIODO " & Combo1.Text
'Text7.Text = I - 1
'MATI12.Row = 1
'MATI12.Col = 1
'MATI12.SetFocus
'Screen.MousePointer = 0
End Sub

Private Sub Command10_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    If Label4.Caption = "" Then
        MsgBox "Seleccione primero el grupo y el área y presione Ok", 48
        Exit Sub
    End If
    If Val(Text7.Text) = 1 Then
        MsgBox "No se puede eliminar el último alumno de la lista", 32, "Eliminar"
        Exit Sub
    End If
    TTT = InputBox("Escriba el código que desea eliminar" & Chr(13) & "(Escriba un número entre 1 y " & Text7.Text & ")", "Eliminar alumno")
    If TTT = "" Then
        MsgBox "No escribió el código", 64, "Eliminar"
        Exit Sub
    End If
    If Val(TTT) > Val(Text7.Text) Or (Val(TTT) < 1) Then
        MsgBox "No existe este código en la lista", 32, "Eliminar"
        Exit Sub
    End If
    MATI12.RemoveItem Val(TTT)
    Text7.Text = Val(Text7.Text) - 1
    For TT = 1 To Val(MATI12.Rows - 1)
        MATI12.TextMatrix(TT, 0) = TT
    Next TT
    VALI4 = False
End If
End Sub

Private Sub Command2_Click()
Dim SumaX As Single
If Val(Text7.Text) = 0 Then
   MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 48, "IMPRIMIR"
   Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
   Screen.MousePointer = 11
   NAR = FreeFile
   Open Ruta & "inicial.edu" For Input As #NAR
   Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
   Close #NAR
   Printer.ScaleMode = 7
   Printer.Font.Size = 10
   Printer.CurrentY = 1
   Printer.CurrentX = 6.5
   Printer.Print "OBSERVACIONES GENERALES " & Frame2.Caption
   Printer.Print ""
   Printer.CurrentX = 0.5
   Printer.Print ini.nombre;
   Printer.CurrentX = 16.5
   Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
   Printer.CurrentX = 0.5
   Printer.Print Frame1.Caption
   Printer.Print ""
   Printer.Print ""
   For I = 0 To (MATI12.Rows - 1)
       If I = 0 Then
            SumaX = 0
      Else
            SumaX = 0.2
      End If
      Printer.CurrentX = 0.5
      'Printer.Print I;
      Printer.Print RTrim(MATI12.TextMatrix(I, 0));
      Printer.CurrentX = 1.3
      Printer.Print RTrim(MATI12.TextMatrix(I, 12));
      Printer.CurrentX = 10.5 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 1));
      Printer.CurrentX = 11.4 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 2));
      Printer.CurrentX = 12.3 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 3));
      Printer.CurrentX = 13.2 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 4));
      Printer.CurrentX = 14.1 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 5));
      Printer.CurrentX = 15 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 6));
      Printer.CurrentX = 15.9 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 7));
      Printer.CurrentX = 16.8 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 8));
      Printer.CurrentX = 17.7 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 9));
      Printer.CurrentX = 18.6 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 10));
      Printer.CurrentX = 19.5 + SumaX
      Printer.Print RTrim(MATI12.TextMatrix(I, 11))
   Next I
   Printer.EndDoc
   Printer.Font.Size = 8
   Screen.MousePointer = 0
End If
End Sub

Private Sub Command3_Click()
Dim colcopy As String
colcopy = InputBox("Código de estudiante a copiar?" & Chr(13) & "(escriba un número de 1 a " & Text7.Text & ")", "copiar")
If colcopy = "" Then Exit Sub
If (Val(colcopy) < 1) Or (Val(colcopy) > Val(Text7.Text)) Then
    MsgBox "Código no existe", 48, "copiar"
    Exit Sub
End If
For kcpy = 0 To 10
    simucopy(kcpy) = MATI12.TextMatrix(Val(colcopy), kcpy + 1)
Next kcpy
Command4.Enabled = True
Command5.Enabled = True
Command8.Enabled = True
End Sub

Private Sub Command4_Click()
For kcpy = 0 To 10
        MATI12.TextMatrix(MATI12.Row, kcpy + 1) = simucopy(kcpy)
Next kcpy
VALI4 = False
End Sub

Private Sub Command5_Click()
RESP = MsgBox("Desea pegar los datos en toda la lista?", vbYesNo + vbQuestion + vbDefaultButton2, "Pegar todo")
If RESP = vbYes Then
    For vcpy = 1 To Val(Text7.Text)
        For kcpy = 0 To 10
            MATI12.TextMatrix(vcpy, kcpy + 1) = simucopy(kcpy)
        Next kcpy
    Next vcpy
End If
VALI4 = False
End Sub

Private Sub Command6_Click()
If Val(Text7.Text) = 0 Then
    MsgBox "ESCOJA EL NOMBRE DEL GRUPO, EL AREA Y PRESIONE OK", 48, "GUARDAR"
    Combo2.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA GUARDAR ESTA INFORMACION PARA EL " & Frame2.Caption & "?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
    Screen.MousePointer = 11
    If Dir(Ruta & Label6.Caption & ".obs") <> "" Then
        Kill Ruta & Label6.Caption & ".obs"
    End If
    VerInfo = False
    For I = 1 To (MATI12.Rows - 1)
        For J = 1 To 11
            If MATI12.TextMatrix(I, J) <> "" Then
                VerInfo = True
            End If
        Next J
    Next I
    If VerInfo = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If
'    NAR = FreeFile
'    Open Ruta & Label6.Caption & ".obs" For Random As #NAR Len = Len(notas)
    For I = 1 To (MATI12.Rows - 1)
        For J = 1 To 10
            If MATI12.TextMatrix(I, J) = "" Then
                notas.area(J) = 0
            Else
                notas.area(J) = RegObserv(MATI12.TextMatrix(I, J))
            End If
        Next J
        'notas.JV = Format(MATI12.TextMatrix(I, 11), ">")
        If MATI12.TextMatrix(I, 11) = "" Then
            notas.FA = 0
        Else
            notas.FA = MATI12.TextMatrix(I, 11)
        End If
        notas.num_carnet = MATI12.TextMatrix(I, 13)
        NAR = FreeFile
        Open Ruta & Label6.Caption & ".obs" For Random As #NAR Len = Len(notas)
        Put #NAR, I, notas
        Close #NAR
    Next I
    'Close #NAR
End If
VALI4 = True
Screen.MousePointer = 0

'If Val(Text7.Text) = 0 Then
'    MsgBox "ESCOJA EL NOMBRE DEL GRUPO, EL AREA Y PRESIONE OK", 48, "GUARDAR"
'    Combo2.SetFocus
'    Exit Sub
'End If
'RESP = MsgBox("DESEA GUARDAR ESTA INFORMACION PARA EL " & Frame2.Caption & "?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
'If RESP = vbYes Then
'    Screen.MousePointer = 11
'    If Dir(Ruta & Label6.Caption & ".obs") <> "" Then
'        Kill Ruta & Label6.Caption & ".obs"
'    End If
'    VerInfo = False
'    For I = 1 To (MATI12.Rows - 1)
'        For J = 1 To 11
'            If MATI12.TextMatrix(I, J) <> "" Then
'                VerInfo = True
'            End If
'        Next J
'    Next I
'    If VerInfo = False Then
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'    NAR = FreeFile
'    Open Ruta & Label6.Caption & ".obs" For Random As #NAR Len = Len(notas)
'    For I = 1 To (MATI12.Rows - 1)
'        For J = 1 To 10
'            If MATI12.TextMatrix(I, J) = "" Then
'                notas.area(J) = 0
'            Else
'                notas.area(J) = MATI12.TextMatrix(I, J)
'            End If
'        Next J
'        'notas.FA = Format(MATI12.TextMatrix(I, 11), ">")
'        If MATI12.TextMatrix(I, 11) = "" Then
'            notas.FA = 0
'        Else
'            notas.FA = MATI12.TextMatrix(I, 11)
'        End If
'        notas.num_carnet = Right(MATI12.TextMatrix(I, 13), 5)
'        Put #NAR, I, notas
'    Next I
'    Close #NAR
'End If
'VALI4 = True
'Screen.MousePointer = 0
End Sub

Private Sub Command7_Click()
SWobserv = True
Ver_Obser.Show
End Sub

Private Sub Command8_Click()
RESP = MsgBox("Desea pegar los datos en toda la columna?", vbYesNo + vbQuestion + vbDefaultButton2, "Pegar todo")
If RESP = vbYes Then
    If MATI12.Col > 0 And MATI12.Col < 11 Then
        For vcpy = 1 To Val(MATI12.Rows - 1)
            MATI12.TextMatrix(vcpy, MATI12.Col) = simucopy(Val(MATI12.Col - 1))
        Next vcpy
    End If
End If
VALI4 = False
End Sub

Private Sub Command9_Click()

I = 0
PASSW.Show 1
If I = 1 Then
    If Label4.Caption = "" Then
        MsgBox "Seleccione primero el grupo y el área y presione Ok", 48
        Exit Sub
    End If
    Screen.MousePointer = 11
    t = 1
    ret = 0
    NAR = FreeFile
    Open Ruta & Label4.Caption & ".gru" For Random As #NAR Len = Len(alugru)
    While Not EOF(NAR)
        ret = ret + 1
        Get #NAR, ret, alugru
    Wend
    Close #NAR
    Open Ruta & Label4.Caption & ".gru" For Random As #NAR Len = Len(alugru)
    For t = 1 To ret - 1
        StudentNew = False
        Get #NAR, t, alugru
        For s = 1 To MATI12.Rows - 1
            If Val(alugru.num_carnet) = Val(MATI12.TextMatrix(s, 13)) Then
                StudentNew = True
                Exit For
            End If
        Next s
        If StudentNew = False Then
            NAR = FreeFile
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, Val(alugru.num_carnet), alumno
            Close #NAR
            NAR = NAR - 1
            Text7.Text = Val(Text7.Text) + 1
            MATI12.Rows = MATI12.Rows + 1
            MATI12.TextMatrix((MATI12.Rows - 1), 12) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
            MATI12.TextMatrix((MATI12.Rows - 1), 13) = alumno.n_carnet
        End If
    Next t
    Close #NAR
    MATI12.Col = 12
    MATI12.Sort = 5
    For TT = 1 To Val(MATI12.Rows - 1)
        MATI12.TextMatrix(TT, 0) = TT
    Next TT
    VALI4 = False
    Screen.MousePointer = 0
End If

'I = 0
'PASSW.Show 1
'If I = 1 Then
'    'Dim alugru As grupoalu
'    'Dim alumno As maestroalum
'    If Label4.Caption = "" Then
'        MsgBox "Seleccione primero el grupo y el área y presione Ok", 48
'        Exit Sub
'    End If
'    TTT = InputBox("Escriba el número de carnet", "Insertar alumno")
'    Screen.MousePointer = 11
'    If TTT = "" Then
'        MsgBox "No escribió el número de carnet", 64, "Insertar"
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'    If Val(TTT) > 32000 Then
'        MsgBox "No. de carnet inválido", 64, "Insertar"
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'    NAR = FreeFile
'    Open Ruta & "cont.edu" For Input As #NAR
'    Input #NAR, I
'    Close #NAR
'    If (Val(TTT) > I - 1) Or (Val(TTT) < 1) Then
'        MsgBox "Registro no existe", 32, "Insertar"
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'    For J = 1 To Val(MATI12.Rows - 1)
'        If Val(Right(MATI12.TextMatrix(J, 13), 5)) = Val(TTT) Then
'            MsgBox "Alumno ya existe en la lista", 32, "Insertar"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'    Next J
'    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'    Get #NAR, Val(TTT), alumno
'    Close #NAR
'    If RTrim(alumno.n_carnet) = "" Then
'        MsgBox "Alumno no existe en base de datos", 32
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'    t = 0
'    s = 0
'    Open Ruta & Label4.Caption & ".gru" For Random As #NAR Len = Len(alugru)
'    While Not EOF(NAR)
'        t = t + 1
'        Get #NAR, t, alugru
'        If Val(alugru.num_carnet) = Val(TTT) Then
'            s = s + 1
'        End If
'    Wend
'    Close #NAR
'    If s = 0 Then
'        MsgBox "Alumno no pertenece a este grupo", 16, "Insertar"
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'    Text7.Text = Val(Text7.Text) + 1
'    MATI12.Rows = MATI12.Rows + 1
'    MATI12.TextMatrix((MATI12.Rows - 1), 12) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
'    MATI12.TextMatrix((MATI12.Rows - 1), 13) = alumno.n_carnet
'    MATI12.Col = 12
'    MATI12.Sort = 5
'    For TT = 1 To Val(MATI12.Rows - 1)
'        MATI12.TextMatrix(TT, 0) = TT
'    Next TT
'    VALI4 = False
'    Screen.MousePointer = 0
'End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Grabación del boletín académico, de acuerdo con el periodo, grupo y área seleccionado."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VALI4 = False Then
   Call Command6_Click
   Unload Me
Else
  Unload Me
End If
Unload Ver_Obser
End Sub

Private Sub MATI12_Click()
Dim ValorObser As Integer

MATI12.ToolTipText = ""
If MATI12.Col > 0 And MATI12.Col < 11 And MATI12.Row > 0 And MATI12.Row <= Val(Text7.Text) Then
   If MATI12.Text = "" Then
      MATI12.ToolTipText = ""
      Exit Sub
   End If
   If Dir(Ruta & fl & ser & que & lw & ".lgr") <> "" Then
      ValorObser = RegObserv(Val(MATI12.Text))
      NAR = FreeFile
      Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
      Get #NAR, ValorObser, logru
      Close #NAR
      MATI12.ToolTipText = "(" & RTrim(logru.indicador) & ") " & RTrim(logru.observ)
   End If
End If

''Dim logru As logris
'MATI12.ToolTipText = ""
'If MATI12.Col > 0 And MATI12.Col < 11 And MATI12.Row > 0 And MATI12.Row <= Val(Text7.Text) Then
'   If MATI12.Text = "" Then
'      MATI12.ToolTipText = ""
'      Exit Sub
'   End If
'   If Dir(Ruta & fl & ser & que & lw & ".lgr") <> "" Then
'      NAR = FreeFile
'      Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
'      Get #NAR, Val(MATI12.Text), logru
'      Close #NAR
'      MATI12.ToolTipText = "(" & RTrim(logru.indicador) & ") " & RTrim(logru.observ)
'   End If
'End If
End Sub

Private Sub MATI12_KeyPress(KeyAscii As Integer)
Dim ValiLgr As Boolean
If MATI12.Col < 12 And MATI12.Col > 0 And MATI12.Row > 0 And MATI12.Row <= Val(Text7.Text) Then
   If KeyAscii = 13 Then
      If MATI12.Col = 11 And MATI12.Row <> Val(Text7.Text) Then
         MATI12.Row = MATI12.Row + 1
         MATI12.Col = 1
         Exit Sub
      End If
      MATI12.Col = MATI12.Col + 1
      Exit Sub
   End If
   C$ = Chr(KeyAscii)
   If KeyAscii = 8 Then
      If MATI12.Text <> "" Then
         MATI12.Text = Left(MATI12.Text, Len(MATI12.Text) - 1)
         If MATI12.Text = "" Then
            MATI12.CellForeColor = RGB(0, 0, 0)
         End If
         VALI4 = False
         Exit Sub
      Else
         If MATI12.Col > 1 Then
            MATI12.Col = MATI12.Col - 1
         End If
      End If
   End If
'   If MATI12.Col = 11 Then
'      If C$ <> "E" And C$ <> "e" And C$ <> "S" And C$ <> "s" And C$ <> "B" And C$ <> "b" And C$ <> "A" And C$ <> "a" And C$ <> "D" And C$ <> "d" And C$ <> "I" And C$ <> "i" Then
'         KeyAscii = 0
'         Beep
'         Exit Sub
'      Else
'         MATI12.CellFontBold = True
'         MATI12.CellForeColor = RGB(255, 0, 0)
'         MATI12.Text = C$
'         VALI4 = False
'         Exit Sub
'      End If
'   End If
   If C$ < "0" Or C$ > "9" Then
      KeyAscii = 0
      Beep
      Exit Sub
   End If
   rete = Chr(KeyAscii)
   If MATI12.Col = 11 Then
      MATI12.CellFontBold = True
      MATI12.CellForeColor = RGB(255, 0, 0)
   Else
      MATI12.CellFontBold = True
      MATI12.CellForeColor = RGB(0, 0, 255)
   End If
   MATI12.Text = MATI12.Text + rete
   VALI4 = False
   If MATI12.Col > 0 And MATI12.Col < 11 Then
      If Val(MATI12.Text) >= CuentaObs Or Val(MATI12.Text) < 1 Then
         MsgBox "OBSERVACION NO EXISTE", 48, "ADVERTENCIA"
         MATI12.Text = ""
         MATI12.CellForeColor = RGB(0, 0, 0)
         VALI4 = False
         'Exit Sub
      End If
   End If
End If

'Dim ValiLgr As Boolean
'If MATI12.Col < 12 And MATI12.Col > 0 And MATI12.Row > 0 And MATI12.Row <= Val(Text7.Text) Then
'   If KeyAscii = 13 Then
'      If MATI12.Col = 11 And MATI12.Row <> Val(Text7.Text) Then
'         MATI12.Row = MATI12.Row + 1
'         MATI12.Col = 1
'         Exit Sub
'      End If
'      MATI12.Col = MATI12.Col + 1
'      Exit Sub
'   End If
'   C$ = Chr(KeyAscii)
'   If KeyAscii = 8 Then
'      If MATI12.Text <> "" Then
'         MATI12.Text = Left(MATI12.Text, Len(MATI12.Text) - 1)
'         If MATI12.Text = "" Then
'            MATI12.CellForeColor = RGB(0, 0, 0)
'         End If
'         VALI4 = False
'         Exit Sub
'      Else
'         If MATI12.Col > 1 Then
'            MATI12.Col = MATI12.Col - 1
'         End If
'      End If
'   End If
''   If MATI12.Col = 11 Then
''      If C$ <> "E" And C$ <> "e" And C$ <> "S" And C$ <> "s" And C$ <> "B" And C$ <> "b" And C$ <> "A" And C$ <> "a" And C$ <> "D" And C$ <> "d" And C$ <> "I" And C$ <> "i" Then
''         KeyAscii = 0
''         Beep
''         Exit Sub
''      Else
''         MATI12.CellFontBold = True
''         MATI12.CellForeColor = RGB(255, 0, 0)
''         MATI12.Text = C$
''         VALI4 = False
''         Exit Sub
''      End If
''   End If
'   If C$ < "0" Or C$ > "9" Then
'      KeyAscii = 0
'      Beep
'      Exit Sub
'   End If
'   rete = Chr(KeyAscii)
'   If MATI12.Col = 11 Then
'      MATI12.CellFontBold = True
'      MATI12.CellForeColor = RGB(255, 0, 0)
'   Else
'      MATI12.CellFontBold = True
'      MATI12.CellForeColor = RGB(0, 0, 255)
'   End If
'   MATI12.Text = MATI12.Text + rete
'   VALI4 = False
'   If MATI12.Col > 0 And MATI12.Col < 11 Then
'      If Val(MATI12.Text) >= FERT Or Val(MATI12.Text) < 1 Then
'         MsgBox "OBSERVACION NO EXISTE", 48, "ADVERTENCIA"
'         MATI12.Text = ""
'         MATI12.CellForeColor = RGB(0, 0, 0)
'         VALI4 = False
'         'Exit Sub
'      End If
'   End If
'End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo3.SetFocus
End If
End Sub

Private Sub Form_Load()
'Dim mate As infomater
'Dim icur As inforcur
MATI12.Row = 0
MATI12.Col = 0
MATI12.ColWidth(0) = 400
MATI12.Text = "CD"
MATI12.Col = 1
MATI12.ColWidth(1) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB1"
MATI12.Col = 2
MATI12.ColWidth(2) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB2"
MATI12.Col = 3
MATI12.ColWidth(3) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB3"
MATI12.Col = 4
MATI12.ColWidth(4) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB4"
MATI12.Col = 5
MATI12.ColWidth(5) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB5"
MATI12.Col = 6
MATI12.ColWidth(6) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB6"
MATI12.Col = 7
MATI12.ColWidth(7) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB7"
MATI12.Col = 8
MATI12.ColWidth(8) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB8"
MATI12.Col = 9
MATI12.ColWidth(9) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB9"
MATI12.Col = 10
MATI12.ColWidth(10) = 500
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = "OB10"
'MATI12.Col = 11
'MATI12.ColWidth(11) = 400
'MATI12.CellForeColor = RGB(255, 255, 255)
'MATI12.CellBackColor = RGB(0, 0, 150)
'MATI12.Text = "J.V."
MATI12.Col = 11
MATI12.ColWidth(11) = 400
MATI12.CellForeColor = RGB(255, 255, 255)
MATI12.CellBackColor = RGB(0, 0, 150)
MATI12.Text = " FA"
MATI12.Col = 12
MATI12.ColWidth(12) = 4200
MATI12.Text = "APELLIDOS Y NOMBRES"
MATI12.Col = 13
MATI12.ColWidth(13) = 1200
MATI12.Text = "No.CARNET"
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
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command2.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Label4.Caption = ""
Label6.Caption = ""
VALI4 = True
End Sub
