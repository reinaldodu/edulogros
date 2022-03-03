VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GRUPOSTEC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de boletín por grado"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "GRUPOSTEC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir"
      Height          =   320
      Left            =   4080
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   9255
      Begin VB.CommandButton Command2 
         Caption         =   "&Copiar"
         Height          =   320
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   320
         Left            =   7800
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "GRUPOSTEC.frx":0442
         Left            =   3000
         List            =   "GRUPOSTEC.frx":0470
         TabIndex        =   2
         Text            =   "PREKINDER"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "GRUPOSTEC.frx":04EC
         Left            =   960
         List            =   "GRUPOSTEC.frx":04FC
         TabIndex        =   1
         Text            =   "UNICA"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "AREA:"
         Height          =   195
         Left            =   4560
         TabIndex        =   13
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "GRADO:"
         Height          =   195
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "GRUPOSTEC.frx":051D
      Left            =   7800
      List            =   "GRUPOSTEC.frx":0530
      TabIndex        =   0
      Text            =   "PRIMERO"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   9255
      Begin MSFlexGridLib.MSFlexGrid MATXECN 
         Height          =   4455
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   1
         Cols            =   16
         FixedCols       =   2
         BackColorBkg    =   12632256
         GridColor       =   12582912
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6960
      TabIndex        =   9
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "GRUPOSTEC"
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
Combo3.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo4.SetFocus
End If
End Sub

Private Sub Combo4_Change()
If Combo4.Text <> Combo4.List(0) Then
    Combo4.Text = Combo4.List(0)
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
PENTI = False
MATXECN.Rows = 1
MATXECN.ToolTipText = ""
Label5.Caption = ""
Y = 0
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    Y = Y + 1
    Get #NAR, Y, mate
    If RTrim(mate.nom) = Combo4.Text Then
        NA = mate.num
    End If
Wend
Close #NAR
cona = 0
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
ser = Left(Combo3.Text, 3)
Label6.Caption = Ruta & fl & Left(Combo3.Text, 3) & NA & lw & ".lgr"
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If (RTrim(icur.jornada) = Combo2.Text) And (RTrim(icur.grado) = Combo3.Text) Then
        If Dir(Ruta & RTrim(icur.nom) & NA & lw & ".obs") <> "" Then
            PENTI = True
            Y = 0
            NAR = FreeFile
            Open Ruta & RTrim(icur.nom) & NA & lw & ".obs" For Random As #NAR Len = Len(notas)
            While Not EOF(NAR)
                Y = Y + 1
                Get #NAR, Y, notas
            Wend
            Close #NAR
            Open Ruta & RTrim(icur.nom) & NA & lw & ".obs" For Random As #NAR Len = Len(notas)
            For I = 1 To (Y - 1)
                Get #NAR, I, notas
                NAR = FreeFile
                Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
                Get #NAR, (Val(notas.num_carnet)), alumno
                Close #NAR
                NAR = NAR - 1
                MATXECN.Rows = MATXECN.Rows + 1
                MATXECN.TextMatrix((MATXECN.Rows - 1), 0) = MATXECN.Rows - 1
                MATXECN.TextMatrix((MATXECN.Rows - 1), 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
                For J = 1 To 10
                    If notas.area(J) = 0 Then
                        MATXECN.TextMatrix((MATXECN.Rows - 1), (J + 1)) = ""
                    Else
                        MATXECN.TextMatrix((MATXECN.Rows - 1), (J + 1)) = notas.area(J)
                    End If
                Next J
                MATXECN.TextMatrix((MATXECN.Rows - 1), 12) = RTrim(notas.FA)
                MATXECN.TextMatrix((MATXECN.Rows - 1), 13) = notas.FA
                MATXECN.TextMatrix((MATXECN.Rows - 1), 14) = alumno.n_carnet
                MATXECN.TextMatrix((MATXECN.Rows - 1), 15) = RTrim(icur.nom)
            Next I
            Close #NAR
            NAR = NAR - 1
        End If
    End If
Wend
Close #NAR
If PENTI = False Then
    MsgBox "No existe información del área " & Format(Combo4.Text, "<") & " para el grado " & Format(Combo3.Text, "<"), 64
Else
    Label5.Caption = "JORNADA:" & Combo2.Text & " - GRADO:" & Combo3.Text & " - AREA:" & Combo4.Text & " - PERIODO:" & Combo1.Text
End If
End Sub

Private Sub Command2_Click()
If MATXECN.Rows = 1 Then
    MsgBox "No existe información para copiar", 64, "Copiar"
    Exit Sub
End If
Clipboard.Clear
cop = ""
For X = 1 To (MATXECN.Rows - 1)
        ape = RTrim(MATXECN.TextMatrix(X, 14))
        nom = RTrim(MATXECN.TextMatrix(X, 1))
        If X < 10 Then
           cop = cop + LTrim(Str(X) & "   - " & ape & "  " & nom) & vbCrLf
        Else
           cop = cop + LTrim(Str(X) & " - " & ape & "  " & nom) & vbCrLf
        End If
Next X
Clipboard.SetText cop
End Sub

Private Sub Command3_Click()
If MATXECN.Rows = 1 Then
    MsgBox "No existe información para imprimir", 64, "Imprimir"
    Exit Sub
End If
IMPARTEC.Show 1
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta de boletín por grado: Para ver la información de la observación, de click en el número de la misma."
End Sub

Private Sub Form_Load()
MATXECN.Row = 0
MATXECN.Col = 0
MATXECN.ColWidth(0) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "CD"
MATXECN.Col = 1
MATXECN.ColWidth(1) = 4200
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "APELLIDOS Y NOMBRES"
MATXECN.Col = 2
MATXECN.ColWidth(2) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB1"
MATXECN.Col = 3
MATXECN.ColWidth(3) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB2"
MATXECN.Col = 4
MATXECN.ColWidth(4) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB3"
MATXECN.Col = 5
MATXECN.ColWidth(5) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB4"
MATXECN.Col = 6
MATXECN.ColWidth(6) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB5"
MATXECN.Col = 7
MATXECN.ColWidth(7) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB6"
MATXECN.Col = 8
MATXECN.ColWidth(8) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB7"
MATXECN.Col = 9
MATXECN.ColWidth(9) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB8"
MATXECN.Col = 10
MATXECN.ColWidth(10) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB9"
MATXECN.Col = 11
MATXECN.ColWidth(11) = 500
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "OB10"
MATXECN.Col = 12
MATXECN.ColWidth(12) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "J.V."
MATXECN.Col = 13
MATXECN.ColWidth(13) = 400
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = " FA"
MATXECN.Col = 14
MATXECN.ColWidth(14) = 1200
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "No.CARNET"
MATXECN.Col = 15
MATXECN.ColWidth(15) = 1600
MATXECN.CellForeColor = RGB(255, 255, 255)
MATXECN.CellBackColor = RGB(0, 0, 150)
MATXECN.Text = "GRUPO"
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    NAR = FreeFile
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    que = 0
    While Not EOF(NAR)
        que = que + 1
        Get #NAR, que, mate
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    For I = 1 To que - 1
        Get #NAR, I, mate
        If RTrim(mate.nom) <> "" Then
            Combo4.AddItem RTrim(mate.nom)
        End If
    Next I
    Close #NAR
    Combo4.Text = Combo4.List(0)
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
Else
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
End If
End Sub

Private Sub MATXECN_Click()
MATXECN.ToolTipText = ""
If MATXECN.Col > 1 And MATXECN.Col < 12 And MATXECN.Row > 0 And MATXECN.Row < (MATXECN.Rows) Then
   If RTrim(MATXECN.Text) = "" Then
      MATXECN.ToolTipText = ""
      Exit Sub
   End If
   If Dir(Label6.Caption) <> "" Then
      NAR = FreeFile
      Open Label6.Caption For Random As #NAR Len = Len(logru)
      Get #NAR, Val(MATXECN.Text), logru
      Close #NAR
      MATXECN.ToolTipText = "(" & RTrim(logru.indicador) & ") " & RTrim(logru.observ)
   End If
End If
End Sub
