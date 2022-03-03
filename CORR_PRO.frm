VERSION 5.00
Begin VB.Form CORR_PRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar datos del profesor"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7830
   Icon            =   "CORR_PRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&GUARDAR"
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      ToolTipText     =   "Guarda la información que se muestra en pantalla"
      Top             =   3960
      Width           =   1815
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
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   320
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Profesor No."
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "-----------------------MODIFICAR DATOS DEL PROFESOR-----------------------"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox Text11 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   " (dd/mm/aaaa)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2760
         TabIndex        =   26
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   240
         TabIndex        =   25
         Top             =   2400
         Width           =   1245
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "TITULO:"
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
         Left            =   3840
         TabIndex        =   22
         Top             =   2880
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ESCALAFON:"
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
         Left            =   5640
         TabIndex        =   21
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label7 
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
         Left            =   3840
         TabIndex        =   20
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TELEFONO:"
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
         Left            =   3840
         TabIndex        =   19
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label Label5 
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
         Left            =   3840
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CEDULA No:"
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
         TabIndex        =   16
         Top             =   1920
         Width           =   1110
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
         Left            =   240
         TabIndex        =   15
         Top             =   1320
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
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   990
      End
   End
End
Attribute VB_Name = "CORR_PRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Luego de haber modificado la información del profesor, de click en Guardar."
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
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
    Call Command1_Click
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
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub
Private Sub TEXT8_KEYPRESS(KeyAscii As Integer)
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
'Dim profe As maestropro
If Text10.Text = "" Then
    MsgBox "ESCRIBA EL NUMERO DEL PROFESOR", 16, "CORREGIR"
    Text10.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
w = Val(Text10.Text)
If ((w > r - 1) Or (w < 1)) Then
    MsgBox "PROFESOR NO EXISTE", 32
    Text10.SetFocus
    VERI = 0
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, w, profe
Close #NAR
If ((RTrim(profe.nombres) = "") And (RTrim(profe.especiali) = "")) Then
    MsgBox "REGISTRO NO EXISTE", 16, "CONSULTAR"
    Text10.SetFocus
    Exit Sub
End If
Text1.Text = RTrim(profe.nombres)
Text2.Text = RTrim(profe.apellidos)
Text3.Text = RTrim(profe.documento)
Text11.Text = RTrim(profe.fech_nacim)
Text4.Text = RTrim(profe.rh)
Text5.Text = RTrim(profe.direccion)
Text6.Text = RTrim(profe.Telefono)
Text7.Text = RTrim(profe.año_ingre)
Text8.Text = RTrim(profe.escalafon)
Text9.Text = RTrim(profe.especiali)
VERI = 1
End Sub
Private Sub Command2_Click()
'Dim profe As maestropro
If VERI = 0 Then
    MsgBox "SELECCIONE PRIMERO EL REGISTRO A CORREGIR", vbCritical, "ADVERTENCIA"
    Text10.SetFocus
    Exit Sub
End If
If (RTrim(Text1.Text) = "") Or (RTrim(Text2.Text) = "") Or (RTrim(Text3.Text) = "") Or (RTrim(Text4.Text) = "") Or (RTrim(Text5.Text) = "") Or (RTrim(Text6.Text) = "") Or (RTrim(Text7.Text) = "") Then
    MsgBox "INFORMACION INCOMPLETA", 32, "GUARDAR"
    Exit Sub
End If
If Val(Text7.Text) < 1900 Then
    MsgBox "AÑO DE INGRESO INCORRECTO", 32, "GUARDAR"
    Text7.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA GUARDAR LA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR REGISTRO")
If RESP = vbYes Then
profe.nombres = Format(Text1.Text, ">")
profe.apellidos = Format(Text2.Text, ">")
profe.documento = Text3.Text
profe.fech_nacim = Text11.Text
profe.rh = Text4.Text
profe.direccion = Text5.Text
profe.Telefono = Text6.Text
profe.año_ingre = Text7.Text
profe.escalafon = Text8.Text
profe.especiali = Format(Text9.Text, ">")
NAR = FreeFile
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Put #NAR, w, profe
Close #NAR
End If
End Sub

Private Sub Form_Load()
VERI = 0
Text1.MaxLength = 20
Text2.MaxLength = 20
Text3.MaxLength = 10
Text4.MaxLength = 4
Text5.MaxLength = 40
Text6.MaxLength = 12
Text7.MaxLength = 4
Text8.MaxLength = 2
Text9.MaxLength = 40
Text10.MaxLength = 3
Text11.MaxLength = 10
End Sub
