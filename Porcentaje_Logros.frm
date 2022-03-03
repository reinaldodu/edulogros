VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Porcentaje_Logros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentajes de logros"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10065
   Icon            =   "Porcentaje_Logros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "PERIODO CUARTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   7680
      TabIndex        =   14
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Ver Logros >>"
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   21
         Top             =   3360
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid Mtx_ptj 
         Height          =   2775
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   30
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   75
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "PERIODO TERCERO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Ver Logros >>"
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   3360
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid Mtx_ptj 
         Height          =   2775
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   1
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   29
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   3000
         Width           =   75
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "PERIODO SEGUNDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   2640
      TabIndex        =   12
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Ver Logros >>"
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Top             =   3360
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid Mtx_ptj 
         Height          =   2775
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   28
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   75
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   5775
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Porcentaje_Logros.frx":0442
         Left            =   1080
         List            =   "Porcentaje_Logros.frx":0452
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Porcentaje_Logros.frx":0473
         Left            =   1080
         List            =   "Porcentaje_Logros.frx":04A1
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MATERIA:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GRADO:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   630
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   8760
      Picture         =   "Porcentaje_Logros.frx":0521
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERIODO PRIMERO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Ver Logros >>"
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   3360
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid Mtx_ptj 
         Height          =   2775
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   27
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   75
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   9120
      TabIndex        =   31
      Top             =   4200
      Width           =   165
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "PORCENTAJE TOTAL..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   6120
      TabIndex        =   26
      Top             =   4200
      Width           =   2880
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6120
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "Porcentaje_Logros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim VerInfo As Boolean, SumPorcent As Integer
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
'VERIFICAMOS SI ESTÁ BLOQUEADO EL PERIODO ACADÉMICO
'If VeriPeriodo(lw) = False Then
'    Exit Sub
'End If

' VERIFICAMOS SI EXISTE INFORMACION PARA GUARDAR
For J = 0 To 3
    If Label1(J) = "Porcentaje..." Then
        VerInfo = False
        For I = 1 To (Mtx_ptj(J).Rows - 1)
            If Trim(Mtx_ptj(J).TextMatrix(I, 1)) <> "" Then
                VerInfo = True
            End If
        Next I
        If VerInfo = True Then
            For I = 1 To (Mtx_ptj(J).Rows - 1)
                If (Trim(Mtx_ptj(J).TextMatrix(I, 1)) = "") Or (Mtx_ptj(J).TextMatrix(I, 1) = "0") Then
                    MsgBox "No ha escrito el porcentaje del logro No." & I & " del periodo " & J + 1, 64, "ADVERTENCIA"
                    Exit Sub
                End If
            Next I
        End If
    End If
Next J

If Val(Label8.Caption) <> 100 And Val(Label8.Caption) <> 0 Then
    MsgBox "La sumatoria de porcentajes debe ser igual a 100", 64, "ADVERTENCIA"
    Exit Sub
End If
  
RESP = MsgBox("DESEA GUARDAR LA INFORMACIÓN DE PORCENTAJES?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
    Screen.MousePointer = 11
    
    For J = 0 To 3
        If Label1(J) = "Porcentaje..." Then
            If Dir(Ruta & Label4.Caption & J + 1 & ".ptj") <> "" Then
                Kill Ruta & Label4.Caption & J + 1 & ".ptj"
            End If
            ' SI NO EXISTE INFORMACION NO SE CREA ARCHIVO .PTJ
            VerInfo = False
            For I = 1 To (Mtx_ptj(J).Rows - 1)
                If Trim(Mtx_ptj(J).TextMatrix(I, 1)) <> "" Then
                    VerInfo = True
                End If
            Next I
            If VerInfo = True Then
                NAR = FreeFile
                Open Ruta & Label4.Caption & J + 1 & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                For I = 1 To (Mtx_ptj(J).Rows - 1)
                    If Trim(Mtx_ptj(J).TextMatrix(I, 1)) = "" Then
                        porcent_manual.porcent_logro = 0
                    Else
                        porcent_manual.porcent_logro = Mtx_ptj(J).TextMatrix(I, 1)
                    End If
                    Put #NAR, I, porcent_manual
                Next I
                Close #NAR
            End If
        End If
    Next J
End If
VALI4 = True
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim ConfLgr As Byte
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
If VALI4 = False Then
    Call Command1_Click
End If
Unload Ver_Obser
Command1.Enabled = False
For J = 0 To 3
    Mtx_ptj(J).Rows = 1
    Mtx_ptj(J).Cols = 2
Next J
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
pio = 0
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If (RTrim(argra.grado) = RTrim(Combo2.Text) And (argra.num_area = que)) Then
        pio = 1
    End If
Wend
Close #NAR
If pio = 0 Then
    MsgBox "NO SE HA CREADO LA MATERIA " & Combo3.Text & " PARA ESTE GRADO", 64, "ADVERTENCIA"
    Combo3.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If

If Combo4.Text = "UNICA" Then
fl = "1"
End If
If Combo4.Text = "MAÑANA" Then
fl = "2"
End If
If Combo4.Text = "TARDE" Then
fl = "3"
End If
If Combo4.Text = "NOCHE" Then
fl = "4"
End If
ser = Left(Combo2.Text, 3)
Label4.Caption = fl & ser & que

Cont_Ttl = 0
For J = 0 To 3
    FERT = 0
    Cont_Lgr = 0
    Open Ruta & fl & ser & que & J + 1 & ".lgr" For Random As #NAR Len = Len(logru)
    While Not EOF(NAR)
        FERT = FERT + 1
        Get #NAR, FERT, logru
        If Trim(logru.indicador) = "L" Then
            Cont_Lgr = Cont_Lgr + 1
        End If
    Wend
    Close #NAR
    If Cont_Lgr = 0 Then
        Label1(J).Caption = "No existen logros"
        Label7(J).Caption = "0"
        Label7(J).Visible = False
        Mtx_ptj(J).Enabled = False
        Command4(J).Enabled = False
    Else
        Label1(J).Caption = "Porcentaje..."
        Mtx_ptj(J).Enabled = True
        Command4(J).Enabled = True
        Cont_Porc_Lgr = 0
        For h = 1 To Cont_Lgr
            Mtx_ptj(J).Rows = Mtx_ptj(J).Rows + 1
            Mtx_ptj(J).TextMatrix(h, 0) = "No." & h
        Next h
        If Dir(Ruta & fl & ser & que & J + 1 & ".ptj") <> "" Then
            Open Ruta & fl & ser & que & J + 1 & ".ptj" For Random As #NAR Len = Len(porcent_manual)
            For h = 1 To Cont_Lgr
                Get #NAR, h, porcent_manual
                If porcent_manual.porcent_logro = 0 Then
                    Mtx_ptj(J).TextMatrix(h, 1) = ""
                Else
                    Mtx_ptj(J).TextMatrix(h, 1) = porcent_manual.porcent_logro
                    Cont_Porc_Lgr = Cont_Porc_Lgr + porcent_manual.porcent_logro
                    Cont_Ttl = Cont_Ttl + porcent_manual.porcent_logro
                End If
            Next h
            Close #NAR
        End If
        Label7(J).Caption = Cont_Porc_Lgr
        Label7(J).Visible = True
    End If
Next J
'If Cont_Lgr > 10 Then
'    MsgBox "NO SE PUEDE CALIFICAR MÁS DE 10 LOGROS POR PERIODO, VERIFIQUE LA CANTIDAD DE LOGROS CREADOS PARA " & Combo3.Text, 64, "ADVERTENCIA"
'    Screen.MousePointer = 0
'    Exit Sub
'End If
Command1.Enabled = True
Label8.Caption = Cont_Ttl

' Se verifica si está bloqueado el periodo para no habilitar el botón guardar
'If VeriPeriodo(lw) = False Then
'    Command1.Enabled = False
'    MsgBox "EL PERIODO... " & " SOLO ESTA DISPONIBLE PARA CONSULTA", 32, "Grabar logros y observaciones"
'Else
'    Command1.Enabled = True
'End If
Frame2.Caption = Combo2.Text & " - " & Combo3.Text
Screen.MousePointer = 0
VALI4 = True
End Sub

Private Sub Command4_Click(index As Integer)
SWobserv = False
lw = index + 1
Ver_Obser.Show
End Sub

Private Sub Form_Load()
Dim ConfLgr As Byte
Command1.Enabled = False
Command2.Enabled = False
For I = 0 To 3
    Mtx_ptj(I).Row = 0
    Mtx_ptj(I).Col = 0
    Mtx_ptj(I).ColWidth(0) = 500
    Mtx_ptj(I).TextMatrix(0, 0) = "Logro"
    Mtx_ptj(I).CellForeColor = RGB(255, 255, 255)
    Mtx_ptj(I).CellBackColor = RGB(0, 0, 150)
    Mtx_ptj(I).Col = 1
    Mtx_ptj(I).ColWidth(1) = 1000
    Mtx_ptj(I).TextMatrix(0, 1) = "Porcentaje"
    Mtx_ptj(I).CellForeColor = RGB(255, 255, 255)
    Mtx_ptj(I).CellBackColor = RGB(0, 0, 150)
    Mtx_ptj(I).Enabled = False
    Command4(I).Enabled = False
    Label7(I).Visible = False
Next I

If Dir(Ruta & "materia.edu") <> "" Then
    Command2.Enabled = True
    NAR = FreeFile
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
    
End If
Combo2.Text = Combo2.List(0)
Combo3.Text = Combo3.List(0)
Combo4.Text = Combo4.List(0)
VALI4 = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VALI4 = False Then
   Call Command1_Click
   Unload Me
Else
  Unload Me
End If
Unload Ver_Obser
End Sub

Private Sub Mtx_ptj_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then
    If Mtx_ptj(index).Text <> "" Then
        Mtx_ptj(index).Text = Left(Mtx_ptj(index).Text, Len(Mtx_ptj(index).Text) - 1)
        VALI4 = False
        GoTo iririr
    End If
End If
C$ = Chr(KeyAscii)
If C$ < "0" Or C$ > "9" Then
      KeyAscii = 0
      Beep
      Exit Sub
End If

rete = C$
Mtx_ptj(index).Text = Mtx_ptj(index).Text + rete
VALI4 = False
If Val(Mtx_ptj(index).Text) > 100 Then
     MsgBox "VALOR DE PORCENTAJE INVÁLIDO", 48, "ADVERTENCIA"
     Mtx_ptj(index).Text = ""
     Exit Sub
End If
iririr:
sumptj = 0
For h = 1 To Mtx_ptj(index).Rows - 1
    sumptj = sumptj + Val(Mtx_ptj(index).TextMatrix(h, 1))
Next h
Label7(index).Caption = sumptj
Label8.Caption = Val(Label7(0).Caption) + Val(Label7(1).Caption) + Val(Label7(2).Caption) + Val(Label7(3).Caption)
End Sub
