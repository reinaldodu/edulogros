VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MATRICULA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matrículas"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   Icon            =   "MATRICULA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&IMPRIMIR"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      ToolTipText     =   "Imprimir la hoja de matrícula del alumno que se muestra en pantalla"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text13 
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
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   6120
      TabIndex        =   40
      Top             =   5400
      Width           =   3135
      Begin VB.CommandButton Command3 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text12 
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
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "CARNET No."
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&MATRICULAR"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      ToolTipText     =   "Matricular al alumno que se muestra en pantalla"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   1000
      Left            =   240
      TabIndex        =   38
      Top             =   5400
      Width           =   3375
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "MATRICULA.frx":0442
         Left            =   960
         List            =   "MATRICULA.frx":0452
         TabIndex        =   53
         Text            =   "UNICA"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   320
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   585
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "MATRICULA.frx":0473
         Left            =   960
         List            =   "MATRICULA.frx":04A4
         TabIndex        =   3
         Text            =   "PREJARDIN"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "GRADO:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9015
      Begin VB.Frame Frame4 
         Caption         =   "COLEGIOS DONDE REALIZÓ ESTUDIOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   55
         Top             =   2880
         Width           =   8775
         Begin MSFlexGridLib.MSFlexGrid Mxmatri 
            Height          =   1455
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   15
            Cols            =   4
            FixedCols       =   0
         End
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text11 
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
         Height          =   315
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "E.P.S:"
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
         Left            =   120
         TabIndex        =   54
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO"
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
         Left            =   6720
         TabIndex        =   51
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "(dd/mm/aaaa)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7080
         TabIndex        =   50
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "EDAD:"
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
         Left            =   7800
         TabIndex        =   49
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "TEL:"
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
         Left            =   2880
         TabIndex        =   48
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "MADRE:"
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
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "TEL:"
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
         Left            =   2880
         TabIndex        =   46
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "PADRE:"
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
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "SEXO:"
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
         Left            =   7800
         TabIndex        =   43
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "GRADO"
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
         Left            =   6720
         TabIndex        =   37
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "AÑO DE INGRESO:"
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
         Left            =   4680
         TabIndex        =   36
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA:"
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
         Left            =   4680
         TabIndex        =   35
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TEL:"
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
         Left            =   2880
         TabIndex        =   34
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DIRECCION:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ACUDIENTE:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.H:"
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
         Left            =   4680
         TabIndex        =   31
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label4 
         Caption         =   "FECHA DE NACIMIENTO:"
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
         Left            =   4440
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DOC. I.D:"
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
         Left            =   4680
         TabIndex        =   29
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "APELLIDOS:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRES:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   990
      End
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "MATRICULAS DE ALUMNOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   840
      TabIndex        =   44
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "CARNET No."
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
      Height          =   195
      Left            =   7080
      TabIndex        =   42
      Top             =   240
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "MATRICULA.frx":052F
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "MATRICULA"
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
    Call Command1_Click
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Command4_Click()
'Dim alumno As maestroalum
Dim FolRow As Byte, FolCol As Byte, CtFol As Byte
Dim NumFol As String, InpFolio As String
If Text13.Text = "" Then
    MsgBox "ESCRIBA PRIMERO EL No. DE CARNET", 16, "IMPRIMIR"
    Text12.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR LA HOJA DE MATRICULA DEL ESTUDIANTE?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR MATRICULA")
If RESP = vbYes Then
    Printer.ScaleMode = 7
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
    Close #NAR
    Printer.Font.Size = 14
    Printer.CurrentY = 1
    Printer.CurrentX = 10.5 - ((Len(ini.nombre) / 3.3) / 2)
    Printer.Print ini.nombre
    Printer.Font.Size = 12
    Printer.CurrentX = 10.2 - ((Len(ini.ciudad) / 4) / 2)
    Printer.Print ini.ciudad
    Printer.CurrentX = 10.2 - ((Len(ini.Rector) / 5.2) / 2)
    Printer.Print ini.Rector
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, (Val(Text13)), alumno
    Close #NAR
    Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
    Get #NAR, (Val(Text13)), AdiCampo
    Close #NAR
    'Encontrar número de folio
    For k = 0 To 13
        If RTrim(alumno.grado) = Combo1.List(k) Then
            If k >= 0 And k < 3 Then FolRow = 1
            If k > 2 And k < 8 Then FolRow = 2
            If k > 7 And k < 14 Then FolRow = 3
            Exit For
        End If
    Next k
    If RTrim(alumno.jornada) = "UNICA" Then FolCol = 1
    If RTrim(alumno.jornada) = "MAÑANA" Then FolCol = 2
    If RTrim(alumno.jornada) = "TARDE" Then FolCol = 3
    If RTrim(alumno.jornada) = "NOCHE" Then FolCol = 4
    If Dir(Ruta & "folio.edu") <> "" Then
        CtFol = 1
        Open Ruta & "folio.edu" For Input As #NAR
        While Not EOF(NAR)
            Input #NAR, InpFolio
            If ((FolRow * 4) - (4 - FolCol)) = CtFol Then
                NumFol = InpFolio
            End If
            CtFol = CtFol + 1
        Wend
        Close #NAR
    End If
    Printer.Font.Size = 10
    Printer.CurrentY = 1
    Printer.CurrentX = 16.5
    Printer.Print "FOLIO No." & NumFol
    Printer.CurrentX = 16.5
    Printer.Print "MATRICULA No."; alumno.n_matricula
    Printer.Line (17, 3)-(20, 3)
    Printer.Line (17, 7)-(20, 7)
    Printer.Line (17, 3)-(17, 7)
    Printer.Line (20, 3)-(20, 7)
    Printer.CurrentY = 4.5
    Printer.CurrentX = 2
    Printer.Print "GRADO: " & alumno.grado
    Printer.CurrentX = 2
    Printer.Print "JORNADA " & alumno.jornada
    Printer.CurrentY = 6
    Printer.CurrentX = 2
    Printer.Print "Fecha de inscripción: " & Format(Date, "mmmm dd/yyyy")
    Printer.CurrentY = 7
    Printer.CurrentX = 2
    Printer.Font.Size = 12
    Printer.Print "1. DATOS PERSONALES"
    Printer.Font.Size = 10
    Printer.Line (2, 7.5)-(20, 7.5)
    Printer.Line (2, 10.2)-(20, 10.2)
    Printer.Line (2, 7.5)-(2, 10.2)
    Printer.Line (20, 7.5)-(20, 10.2)
    Printer.CurrentY = 7.8
    Printer.CurrentX = 2.3
    Printer.Print "APELLIDOS: " & alumno.apellidos
    Printer.CurrentX = 2.3
    Printer.Print "NOMBRES: " & alumno.nombres
    Printer.CurrentX = 2.3
    Printer.Print "DOCUMENTO DE IDENTIDAD: " & alumno.documento
    Printer.CurrentX = 2.3
    Printer.Print "FECHA DE NACIMIENTO: " & alumno.f_nacimiento;
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
    Printer.CurrentX = 15
    Printer.Print "EDAD ACTUAL: " & aaaa
    Printer.CurrentX = 2.3
    Printer.Print "DIRECCION DE LA RESIDENCIA: " & alumno.direccion;
    Printer.CurrentX = 15
    Printer.Print "TELEFONO: " & AdiCampo.Tel_casa
    Printer.CurrentY = 10.5
    Printer.Font.Size = 12
    Printer.CurrentX = 2
    Printer.Print "2. COMPOSICIÓN FAMILIAR"
    Printer.Font.Size = 10
    Printer.Line (2, 11)-(20, 11)
    Printer.Line (2, 13.7)-(20, 13.7)
    Printer.Line (2, 11)-(2, 13.7)
    Printer.Line (20, 11)-(20, 13.7)
    Printer.CurrentY = 11.2
    Printer.CurrentX = 8
    Printer.Print "INFORMACION ACERCA DEL PADRE"
    Printer.Line (2, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.CurrentX = 2.3
    Printer.Print "NOMBRES Y APELLIDOS: "; alumno.padre;
    Printer.CurrentX = 15.5
    Printer.Print "TELEFONO: " & alumno.tel_pa
    Printer.Line (2, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.CurrentX = 8
    Printer.Print "INFORMACION ACERCA DE LA MADRE"
    Printer.Line (2, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.CurrentX = 2.3
    Printer.Print "NOMBRES Y APELLIDOS: "; alumno.madre;
    Printer.CurrentX = 15.5
    Printer.Print "TELEFONO: " & alumno.tel_ma
    Printer.CurrentY = 14
    Printer.CurrentX = 2
    Printer.Font.Size = 12
    Printer.Print "3. COLEGIO DONDE REALIZÓ ESTUDIOS"
    Printer.Font.Size = 10
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Line (2, Printer.CurrentY)-(20, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.1
    For k = 0 To 14
        If Mxmatri.TextMatrix(k, 0) <> "" Then
            Printer.CurrentX = 2
            Printer.Print Mxmatri.TextMatrix(k, 0);
            Printer.CurrentX = 10
            Printer.Print Mxmatri.TextMatrix(k, 1);
            Printer.CurrentX = 14.5
            Printer.Print Mxmatri.TextMatrix(k, 2);
            Printer.CurrentX = 16.5
            Printer.Print Mxmatri.TextMatrix(k, 3)
            If k = 0 Then
                Printer.CurrentY = Printer.CurrentY + 0.1
                Printer.Line (2, Printer.CurrentY)-(20, Printer.CurrentY)
                Printer.CurrentY = Printer.CurrentY + 0.1
            End If
        End If
    Next k
    Printer.Line (2, 21)-(7, 21)
    Printer.Line (13, 21)-(18, 21)
    Printer.CurrentY = 21.1
    Printer.CurrentX = 3
    Printer.Print "Padre o acudiente";
    Printer.CurrentX = 14
    Printer.Print "Firma del estudiante"
    Printer.Line (2, 23)-(7, 23)
    Printer.Line (13, 23)-(18, 23)
    Printer.CurrentY = 23.1
    Printer.CurrentX = 3
    Printer.Print vini.VRector;
    Printer.CurrentX = 14.5
    Printer.Print "Secretaria(o)"
    Printer.CurrentY = 24
    Printer.CurrentX = 2
    Printer.Print "OBSERVACIONES:"
    Printer.Line (6, Printer.CurrentY)-(20, Printer.CurrentY)
    Printer.Line (2, Printer.CurrentY + 0.5)-(20, Printer.CurrentY + 0.5)
    Printer.Line (2, Printer.CurrentY + 0.5)-(20, Printer.CurrentY + 0.5)
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Para matricular un alumno seleccione primero la jornada y el grado, y luego de click en Ok."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VERIFI2 = False Then
   Call Command2_Click
   Unload Me
Else
   Unload Me
End If
End Sub

Private Sub Mxmatri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Mxmatri.Col = 3 And Mxmatri.Row <> 14 Then
        Mxmatri.Row = Mxmatri.Row + 1
        Mxmatri.Col = 0
        Exit Sub
    End If
    If Mxmatri.Row = 14 And Mxmatri.Col = 3 Then
    Exit Sub
    Else
        Mxmatri.Col = Mxmatri.Col + 1
        Exit Sub
    End If
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
   If Mxmatri.Text <> "" Then
      Mxmatri.Text = Left(Mxmatri.Text, Len(Mxmatri.Text) - 1)
      VERIFI2 = False
      Exit Sub
   Else
      If Mxmatri.Col <> 0 Then
         Mxmatri.Col = Mxmatri.Col - 1
         Exit Sub
      End If
   End If
   Exit Sub
End If
Mxmatri.Text = Mxmatri.Text + Chr(KeyAscii)
VERIFI2 = False
End Sub

Private Sub TEXT12_KEYPRESS(KeyAscii As Integer)
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
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "NO HAY INFORMACION PARA MATRICULAR", 16, "ADVERTENCIA"
    Text12.SetFocus
    Exit Sub
End If
If (Text20.Text <> "SIN GRUPO") And (Text20.Text <> "PENDIENTE") Then
    MsgBox "NO SE PUEDE MATRICULAR A UN ESTUDIANTE QUE PERTENEZCA A UN GRUPO", 64, "MATRICULAR"
    Text12.SetFocus
    Exit Sub
End If
Text11.Text = Combo1.Text
Text9.Text = Combo2.Text
VERIFI2 = False
Command2.SetFocus
End Sub

Private Sub Command2_Click()
'Dim alumno As maestroalum
If Text1.Text = "" Then
    MsgBox "ESCRIBA PRIMERO EL NUMERO DE CARNET PARA MATRICULAR", 32, "ADVERTENCIA"
    Text12.SetFocus
    Exit Sub
End If
If (Text20.Text <> "SIN GRUPO") And (Text20.Text <> "PENDIENTE") Then
    MsgBox "NO SE PUEDE MATRICULAR A UN ESTUDIANTE QUE PERTENEZCA A UN GRUPO", 64, "MATRICULAR"
    Text12.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA MATRICULARLO(A) PARA ESTE GRADO?", vbYesNo + vbQuestion + vbDefaultButton1, "Carnet No." & Text13.Text)
If RESP = vbYes Then
    AT = Text9.Text
    If AT = "UNICA" Then
        JO = "1"
    End If
    If AT = "MAÑANA" Then
        JO = "2"
    End If
    If AT = "TARDE" Then
        JO = "3"
    End If
    If AT = "NOCHE" Then
        JO = "4"
    End If
    If alumno.n_matricula = 0 Then
        Open Ruta & "conmatri.edu" For Input As #NAR
        Input #NAR, zo
        Close #NAR
        Frame1.Caption = "MATRICULA No." & zo
        alumno.n_matricula = zo
        zo = zo + 1
        Open Ruta & "conmatri.edu" For Output As #NAR
        Print #NAR, zo
        Close #NAR
    End If
    alumno.n_carnet = Text13.Text
    alumno.nombres = Text1.Text
    alumno.apellidos = Text2.Text
    alumno.documento = Text3.Text
    alumno.f_nacimiento = Text4.Text
    alumno.rh = Text5.Text
    alumno.acudiente = Text6.Text
    alumno.tel_acu = Text8.Text
    alumno.padre = Text15.Text
    alumno.tel_pa = Text16.Text
    alumno.madre = Text17.Text
    alumno.tel_ma = Text18.Text
    alumno.direccion = Text7.Text
    alumno.jornada = Text9.Text
    alumno.año_ingre = Text10.Text
    alumno.grado = Text11.Text
    alumno.sexo = Text14.Text
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Put #NAR, h, alumno
    Close #NAR
    Open Ruta & "infcol.edu" For Random As #NAR Len = Len(newmatri)
    For J = 1 To 14
        newmatri.nombre(J) = Mxmatri.TextMatrix(J, 0)
        newmatri.grado(J) = Mxmatri.TextMatrix(J, 1)
        newmatri.año(J) = Mxmatri.TextMatrix(J, 2)
        newmatri.ciudad(J) = Mxmatri.TextMatrix(J, 3)
    Next J
    Put #NAR, h, newmatri
    Close #NAR
End If
VERIFI2 = True
Text12.SetFocus
End Sub

Private Sub Command3_Click()
'Dim alumno As maestroalum
'Dim aluper As pertgrup
If VERIFI2 = False Then
    Call Command2_Click
End If
If Text12.Text = "" Then
    MsgBox "ESCRIBA UN NUMERO DE CARNET", 48, "MATRICULA"
    Text12.SetFocus
    Exit Sub
End If
If Val(Text12.Text) > 32000 Then
    MsgBox "No. DE CARNET INVALIDO", 48, "MATRICULA"
    Text12.SetFocus
    Exit Sub
End If
Frame1.Caption = ""
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
h = Val(Text12.Text)
If ((h > I - 1) Or (h < 1)) Then
    MsgBox "REGISTRO NO EXISTE", 64, "MATRICULA"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
    Text12.SetFocus
    Exit Sub
End If
'Reinicia archivo de informacion de colegios sino existe
If Dir(Ruta & "infcol.edu") = "" Then
    Screen.MousePointer = 11
    Open Ruta & "infcol.edu" For Random As #NAR Len = Len(newmatri)
    For J = 1 To (I - 1)
        For Y = 1 To 14
            newmatri.nombre(Y) = ""
            newmatri.grado(Y) = ""
            newmatri.año(Y) = ""
            newmatri.ciudad(Y) = ""
        Next Y
        Put #NAR, J, newmatri
    Next J
    Close #NAR
    Screen.MousePointer = 0
End If
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
    MsgBox "REGISTRO NO EXISTE", 32, "MATRICULA"
    Text12.SetFocus
    Exit Sub
End If
'Muestra informacion de colegios
Open Ruta & "infcol.edu" For Random As #NAR Len = Len(newmatri)
Get #NAR, h, newmatri
For J = 1 To 14
    Mxmatri.TextMatrix(J, 0) = RTrim(Format(newmatri.nombre(J), ">"))
    Mxmatri.TextMatrix(J, 1) = RTrim(Format(newmatri.grado(J), ">"))
    Mxmatri.TextMatrix(J, 2) = RTrim(newmatri.año(J))
    Mxmatri.TextMatrix(J, 3) = RTrim(Format(newmatri.ciudad(J), ">"))
Next J
Close #NAR
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
Frame1.Caption = "MATRICULA No." & alumno.n_matricula
Text1.Text = RTrim(alumno.nombres)
Text2.Text = RTrim(alumno.apellidos)
Text3.Text = RTrim(alumno.documento)
Text4.Text = RTrim(alumno.f_nacimiento)
Text5.Text = RTrim(alumno.rh)
Text6.Text = RTrim(alumno.acudiente)
Text8.Text = RTrim(alumno.tel_acu)
Text15.Text = RTrim(alumno.padre)
Text16.Text = RTrim(alumno.tel_pa)
Text17.Text = RTrim(alumno.madre)
Text18.Text = RTrim(alumno.tel_ma)
Text7.Text = RTrim(alumno.direccion)
Text9.Text = RTrim(alumno.jornada)
Text10.Text = RTrim(alumno.año_ingre)
Text11.Text = RTrim(alumno.grado)
Text13.Text = RTrim(alumno.n_carnet)
Text14.Text = RTrim(alumno.sexo)
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
Text19.Text = aaaa
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
Get #NAR, h, aluper
Close #NAR
Text20.Text = RTrim(aluper.grupo)
Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
Get #NAR, h, AdiCampo
Close #NAR
Text21.Text = RTrim(AdiCampo.salud)
Text12.Text = ""
Combo1.SetFocus
End Sub

Private Sub Form_Load()
Text12.MaxLength = 5
VERIFI2 = True
Mxmatri.ColWidth(0) = 3600
Mxmatri.TextMatrix(0, 0) = "NOMBRE"
Mxmatri.ColWidth(1) = 2000
Mxmatri.TextMatrix(0, 1) = "GRADO"
Mxmatri.ColWidth(2) = 1000
Mxmatri.TextMatrix(0, 2) = "AÑO"
Mxmatri.ColWidth(3) = 1600
Mxmatri.TextMatrix(0, 3) = "CIUDAD"
End Sub
