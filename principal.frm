VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BASE_ALUM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base de datos de estudiantes"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   Icon            =   "principal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
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
      Height          =   320
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&MODIFICAR"
      Height          =   855
      Left            =   2160
      MaskColor       =   &H0000FFFF&
      Picture         =   "principal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Modifica la información de un alumno existente"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONSULTAR ESTUDIANTE"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   3960
      TabIndex        =   12
      Top             =   5520
      Width           =   5295
      Begin VB.CommandButton Command2 
         Caption         =   "&Ok"
         Height          =   300
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text10 
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
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "(Digite los últimos cinco (5) números del carnet)."
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3120
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "CARNET No."
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&GUARDAR"
      Height          =   855
      Left            =   240
      Picture         =   "principal.frx":0864
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Guarda la información que se encuentra en pantalla"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   " ------------------------------------------- ESTUDIANTE NUEVO ---------------------------------------"
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
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   9015
      Begin VB.CommandButton Command4 
         Caption         =   "Importar CSV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   50
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text18 
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
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox Text17 
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
         Height          =   285
         Left            =   1440
         TabIndex        =   30
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Text9 
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
         Height          =   285
         Left            =   6360
         TabIndex        =   41
         Top             =   3120
         Width           =   2415
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3720
         TabIndex        =   27
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3720
         TabIndex        =   25
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   26
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   24
         Top             =   1320
         Width           =   1575
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7080
         TabIndex        =   35
         ToolTipText     =   "Año"
         Top             =   600
         Width           =   550
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
         Left            =   6720
         TabIndex        =   34
         ToolTipText     =   "Mes"
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "principal.frx":0C42
         Left            =   6360
         List            =   "principal.frx":0C52
         TabIndex        =   39
         Text            =   "UNICA"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3720
         TabIndex        =   29
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text7 
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
         Left            =   1440
         TabIndex        =   31
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "principal.frx":0C73
         Left            =   8160
         List            =   "principal.frx":0C7D
         TabIndex        =   37
         Text            =   "M"
         Top             =   1080
         Width           =   615
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "principal.frx":0C87
         Left            =   6360
         List            =   "principal.frx":0CA3
         TabIndex        =   36
         Text            =   "A +"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text5 
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
         Left            =   6360
         TabIndex        =   33
         ToolTipText     =   "Día"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text4 
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
         Left            =   6360
         TabIndex        =   38
         Top             =   1560
         Width           =   1695
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
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text2 
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
         Left            =   1440
         TabIndex        =   22
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL:"
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
         Left            =   240
         TabIndex        =   49
         Top             =   3240
         Width           =   630
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "TEL. CASA:"
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
         Left            =   240
         TabIndex        =   48
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label12 
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
         Left            =   5040
         TabIndex        =   47
         Top             =   3240
         Width           =   555
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
         Left            =   3120
         TabIndex        =   46
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label19 
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
         Left            =   3120
         TabIndex        =   45
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label18 
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
         Left            =   240
         TabIndex        =   44
         Top             =   1800
         Width           =   735
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
         Left            =   240
         TabIndex        =   43
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label16 
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
         Left            =   7440
         TabIndex        =   21
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label15 
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
         Left            =   5040
         TabIndex        =   20
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "(dd/mm/aaaa)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7680
         TabIndex        =   11
         Top             =   720
         Width           =   1020
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
         Left            =   5040
         TabIndex        =   10
         Top             =   2520
         Width           =   930
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
         Left            =   5040
         TabIndex        =   9
         Top             =   2160
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
         Left            =   3120
         TabIndex        =   8
         Top             =   2160
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
         Left            =   240
         TabIndex        =   7
         Top             =   2880
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
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "FACTOR R-H:"
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
         Left            =   5040
         TabIndex        =   5
         Top             =   1200
         Width           =   1200
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
         Left            =   5040
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00400040&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   990
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MATI1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      BackColor       =   16777215
      ForeColor       =   0
      BackColorBkg    =   -2147483633
      GridColor       =   12582912
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ESTUDIANTES EXISTENTES..."
      Height          =   195
      Left            =   6120
      TabIndex        =   18
      Top             =   6480
      Width           =   2325
   End
End
Attribute VB_Name = "BASE_ALUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Text4.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.SetFocus
End If
End Sub

Private Sub Command4_Click()
I = 0
PASSW.Show 1
If I = 1 Then
Import_CSV.Show
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Agrega información de alumnos a la base de datos principal."
End Sub

Private Sub Text11_Change()
If Len(Text11.Text) = 2 Then
Text12.SetFocus
End If
If Len(Text11.Text) = 0 Then
Text5.SetFocus
End If
End Sub

Private Sub TEXT11_KEYPRESS(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text11.Text) = 1 Then
        Text11.Text = "0" & Text11.Text
    End If
    Text12.SetFocus
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

Private Sub Text12_Change()
If Len(Text12.Text) = 4 Then
Combo1.SetFocus
End If
If Len(Text12.Text) = 0 Then
Text11.SetFocus
End If
End Sub

Private Sub TEXT12_KEYPRESS(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
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

Private Sub TEXT13_KEYPRESS(KeyAscii As Integer)
If KeyAscii = 13 Then
Text15.SetFocus
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text16.SetFocus
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text14.SetFocus
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

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6.SetFocus
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

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7.SetFocus
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

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text13.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo4.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
'If C$ < "0" Or C$ > "9" Then
'    KeyAscii = 0
'    Beep
'End If
End Sub

Private Sub Text5_Change()
If Len(Text5.Text) = 2 Then
Text11.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text5.Text) = 1 Then
        Text5.Text = "0" & Text5.Text
    End If
    Text11.SetFocus
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text8.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text18.SetFocus
End If
End Sub

Private Sub TEXT8_KEYPRESS(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text17.SetFocus
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
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command2_Click
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
'Dim alumno As maestroalum
'Dim aluper As pertgrup
'Dim pens(1 To 12) As Currency
If (RTrim(Text2.Text) = "") Or (RTrim(Text3.Text) = "") Or (RTrim(Text7.Text) = "") Or (RTrim(Text5.Text) = "") Or (RTrim(Text11.Text) = "") Or (RTrim(Text12.Text) = "") Then
    MsgBox "INFORMACION INCOMPLETA", 16, "BASE DE ESTUDIANTES"
    Exit Sub
End If
If (Val(Text5.Text) < 1) Or (Val(Text5.Text) > 31) Then
    MsgBox "DIA INVALIDO", 48, "FECHA DE NACIMIENTO"
    Text5.SetFocus
    Exit Sub
End If
If (Val(Text11.Text) < 1) Or (Val(Text11.Text) > 12) Then
    MsgBox "MES INVALIDO", 48, "FECHA DE NACIMIENTO"
    Text11.SetFocus
    Exit Sub
End If
If Val(Text12.Text) < 1900 Then
    MsgBox "AÑO INVALIDO", 48, "FECHA DE NACIMIENTO"
    Text12.SetFocus
    Exit Sub
End If
If Len(Text5.Text) = 1 Then
Text5.Text = "0" & Text5.Text
End If
If Len(Text11.Text) = 1 Then
Text11.Text = "0" & Text11.Text
End If
sir = 0
rei = 0
NAR = FreeFile
Open Ruta & "infcaret.edu" For Random As #NAR Len = 2
While Not EOF(NAR)
sir = sir + 1
Get #NAR, sir, clat
If clat <> 0 Then
I = clat
clat = 0
Put #NAR, sir, clat
Close #NAR
rei = 1
GoTo oto
End If
Wend
Close #NAR
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
oto:
If I < 10 Then
CER = "0000"
End If
If ((I > 9) And (I < 100)) Then
CER = "000"
End If
If ((I > 99) And (I < 1000)) Then
CER = "00"
End If
If ((I > 999) And (I < 10000)) Then
CER = "0"
End If
If ((I > 9999) And (I < 100000)) Then
CER = ""
End If
alumno.n_matricula = 0
alumno.grado = "SIN GRADO"
alumno.nombres = Format(Text2.Text, ">")
alumno.apellidos = Format(Text3.Text, ">")
alumno.documento = Text4.Text
alumno.f_nacimiento = Text5.Text & "/" & Text11.Text & "/" & Text12.Text
alumno.rh = Combo1.Text
alumno.sexo = Combo2.Text
alumno.padre = Format(Text13.Text, ">")
alumno.tel_pa = Text15.Text
alumno.madre = Format(Text14.Text, ">")
alumno.tel_ma = Text16.Text
alumno.acudiente = Format(Text6.Text, ">")
alumno.tel_acu = Text8.Text
alumno.direccion = Text7.Text
alumno.jornada = Combo3.Text
If Combo3.Text = "UNICA" Then
JO = "1"
End If
If Combo3.Text = "MAÑANA" Then
JO = "2"
End If
If Combo3.Text = "TARDE" Then
JO = "3"
End If
If Combo3.Text = "NOCHE" Then
JO = "4"
End If
alumno.año_ingre = Combo4.Text
AI = Right(Combo4.Text, 2)
alumno.n_carnet = I
MATI1.Rows = MATI1.Rows + 1
MATI1.TextMatrix((MATI1.Rows - 1), 0) = alumno.n_carnet
MATI1.TextMatrix((MATI1.Rows - 1), 1) = RTrim(alumno.nombres)
MATI1.TextMatrix((MATI1.Rows - 1), 2) = RTrim(alumno.apellidos)
MATI1.TextMatrix((MATI1.Rows - 1), 3) = RTrim(alumno.documento)
MATI1.TextMatrix((MATI1.Rows - 1), 4) = RTrim(alumno.f_nacimiento)
MATI1.TextMatrix((MATI1.Rows - 1), 5) = RTrim(alumno.rh)
MATI1.TextMatrix((MATI1.Rows - 1), 6) = RTrim(alumno.acudiente)
MATI1.TextMatrix((MATI1.Rows - 1), 7) = RTrim(alumno.direccion)
MATI1.TextMatrix((MATI1.Rows - 1), 8) = RTrim(alumno.tel_acu)
MATI1.TextMatrix((MATI1.Rows - 1), 9) = RTrim(alumno.jornada)
MATI1.TextMatrix((MATI1.Rows - 1), 10) = RTrim(alumno.año_ingre)
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Put #NAR, I, alumno
Close #NAR
aluper.grupo = "PENDIENTE"
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
Put #NAR, I, aluper
Close #NAR
AdiCampo.otras = ""
AdiCampo.salud = Format(Text9.Text, ">")
AdiCampo.Tel_casa = Text17.Text
AdiCampo.email = Text18.Text
Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
Put #NAR, I, AdiCampo
Close #NAR
If Dir(Ruta & "pensi.edu") <> "" Then
    For J = 1 To 12
        pens(J) = 0
    Next J
    Open Ruta & "pensi.edu" For Random As #NAR Len = 96
    Put #NAR, I, pens
    Close #NAR
End If
If Dir(Ruta & "infcol.edu") <> "" Then
    For J = 1 To 14
        newmatri.nombre(J) = ""
        newmatri.grado(J) = ""
        newmatri.año(J) = ""
        newmatri.ciudad(J) = ""
    Next J
    Open Ruta & "infcol.edu" For Random As #NAR Len = Len(newmatri)
    Put #NAR, I, newmatri
    Close #NAR
End If
If rei = 0 Then
    I = I + 1
    Open Ruta & "cont.edu" For Output As #NAR
    Print #NAR, I
    Close #NAR
End If
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text2.SetFocus
Text1.Text = Text1.Text + 1
End Sub

Private Sub Command2_Click()
'Dim alumno As maestroalum
'Dim aluper As pertgrup
If Text10.Text = "" Then
    MsgBox "ESCRIBA UN NUMERO DE CARNET", 64, "ADVERTENCIA"
    Text10.SetFocus
    Exit Sub
End If
If Val(Text10.Text) > 32000 Then
    MsgBox "No.CARNET INVALIDO", 64, "ADVERTENCIA"
    Text10.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
h = Val(Text10.Text)
If ((h > I - 1) Or (h < 1)) Then
    MsgBox "REGISTRO NO EXISTE", 32
    Text10.SetFocus
    Exit Sub
End If
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
    MsgBox "REGISTRO NO EXISTE", 32
    Text10.SetFocus
    Exit Sub
End If
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
Get #NAR, h, aluper
Close #NAR
Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
Get #NAR, h, AdiCampo
Close #NAR
CONS_ALUM.Text21.Text = RTrim(AdiCampo.salud)
CONS_ALUM.Text1.Text = alumno.n_carnet
CONS_ALUM.Text13.Text = alumno.n_matricula
CONS_ALUM.Text2.Text = RTrim(alumno.nombres)
CONS_ALUM.Text3.Text = RTrim(alumno.apellidos)
CONS_ALUM.Text11.Text = RTrim(alumno.documento)
CONS_ALUM.Text4.Text = RTrim(alumno.f_nacimiento)
CONS_ALUM.Text5.Text = RTrim(alumno.rh)
CONS_ALUM.Text6.Text = RTrim(alumno.acudiente)
CONS_ALUM.Text8.Text = RTrim(alumno.tel_acu)
CONS_ALUM.Text16.Text = RTrim(alumno.padre)
CONS_ALUM.Text17.Text = RTrim(alumno.tel_pa)
CONS_ALUM.Text18.Text = RTrim(alumno.madre)
CONS_ALUM.Text19.Text = RTrim(alumno.tel_ma)
CONS_ALUM.Text22.Text = RTrim(AdiCampo.Tel_casa)
CONS_ALUM.Text23.Text = RTrim(AdiCampo.email)
CONS_ALUM.Text7.Text = RTrim(alumno.direccion)
CONS_ALUM.Text9.Text = RTrim(alumno.jornada)
CONS_ALUM.Text10.Text = RTrim(alumno.año_ingre)
CONS_ALUM.Text12.Text = RTrim(alumno.grado)
CONS_ALUM.Text20.Text = RTrim(aluper.grupo)
CONS_ALUM.Text14.Text = RTrim(alumno.sexo)
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
CONS_ALUM.Text15.Text = aaaa
If Dir(Ruta & "FOTOALU\" & h & ".jpg") <> "" Then
CONS_ALUM.Picture1.Picture = LoadPicture(Ruta & "FOTOALU\" & h & ".jpg")
End If
CONS_ALUM.Show
End Sub

Private Sub Command3_Click()
CORR_ALUM.Show 1
End Sub
Private Sub Form_Load()
Text2.MaxLength = 30
Text3.MaxLength = 30
Text4.MaxLength = 15
Text5.MaxLength = 2
Text6.MaxLength = 30
Text7.MaxLength = 40
Text8.MaxLength = 12
Text9.MaxLength = 20
Text10.MaxLength = 5
Text11.MaxLength = 2
Text12.MaxLength = 4
Text13.MaxLength = 30
Text14.MaxLength = 30
Text15.MaxLength = 15
Text16.MaxLength = 15
Text17.MaxLength = 15
Text18.MaxLength = 50
MATI1.Row = 0
MATI1.Col = 0
MATI1.ColWidth(0) = 1200
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "No.CARNET"
MATI1.Col = 1
MATI1.ColWidth(1) = 2500
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "NOMBRES"
MATI1.Col = 2
MATI1.ColWidth(2) = 2500
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "APELLIDOS"
MATI1.Col = 3
MATI1.ColWidth(3) = 1200
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "DOC. I.D"
MATI1.Col = 4
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "F_NACIM"
MATI1.Col = 5
MATI1.ColWidth(5) = 500
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "R-H"
MATI1.Col = 6
MATI1.ColWidth(6) = 3000
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "ACUDIENTE"
MATI1.Col = 7
MATI1.ColWidth(7) = 4000
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "DIRECCION"
MATI1.Col = 8
MATI1.ColWidth(8) = 1200
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "TELEFONO"
MATI1.Col = 9
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "JORNADA"
MATI1.Col = 10
MATI1.CellForeColor = RGB(255, 255, 255)
MATI1.CellBackColor = RGB(0, 0, 150)
MATI1.Text = "INGRESO"
For J = 2000 To 2100
Combo4.AddItem J
Next J
'Combo4.Text = Combo4.List(0)
Combo4.Text = Combo4.List(Right(Year(Date), 3))
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
sir = 0
SIRO = 0
Open Ruta & "infcaret.edu" For Random As #NAR Len = 2
While Not EOF(NAR)
sir = sir + 1
Get #NAR, sir, clat
If clat <> 0 Then
SIRO = SIRO + 1
End If
Wend
Close #NAR
Text1.Text = (I - 1) - SIRO
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub
