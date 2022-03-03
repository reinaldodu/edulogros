VERSION 5.00
Begin VB.Form BASE_PROF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base de datos de profesores"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9375
   Icon            =   "BASE_PROF.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "&LISTA DE PROFESORES"
      Height          =   495
      Left            =   2520
      TabIndex        =   24
      ToolTipText     =   "Muestra la lista de profesores existentes"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&BORRAR PROFESOR"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      ToolTipText     =   "Retirar profesor de la base de datos de profesores existentes"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONSULTAR PROFESOR:"
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   4680
      TabIndex        =   29
      Top             =   3840
      Width           =   4455
      Begin VB.CommandButton Command3 
         Caption         =   "&Aceptar"
         Height          =   300
         Left            =   2040
         TabIndex        =   26
         Top             =   480
         Width           =   1335
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
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1080
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "NUMERO:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.TextBox Text8 
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
      Height          =   320
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&MODIFICAR"
      Height          =   495
      Left            =   2520
      TabIndex        =   22
      ToolTipText     =   "Modifica la información de un profesor existente"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&GUARDAR"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      ToolTipText     =   "Guarda la información que se encuentra en pantalla"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "--------------------------------------------- PROFESOR NUEVO ----------------------------------------"
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
      TabIndex        =   0
      Top             =   120
      Width           =   8895
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2160
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         ToolTipText     =   "Año"
         Top             =   2400
         Width           =   615
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
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         ToolTipText     =   "Mes"
         Top             =   2400
         Width           =   375
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         ToolTipText     =   "Día"
         Top             =   2400
         Width           =   375
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
         ItemData        =   "BASE_PROF.frx":0ABA
         Left            =   8040
         List            =   "BASE_PROF.frx":0AFA
         TabIndex        =   19
         Text            =   "1"
         Top             =   2160
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
         ItemData        =   "BASE_PROF.frx":0B45
         Left            =   1440
         List            =   "BASE_PROF.frx":0B61
         TabIndex        =   15
         Text            =   "O +"
         Top             =   3000
         Width           =   855
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
         Height          =   375
         Left            =   5280
         TabIndex        =   20
         Top             =   2880
         Width           =   3375
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
         Height          =   375
         Left            =   5280
         TabIndex        =   17
         Top             =   1320
         Width           =   1575
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
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   600
         Width           =   3375
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
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
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
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text1 
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
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   " (dd/mm/aaaa)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2760
         TabIndex        =   33
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label Label12 
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
         Left            =   120
         TabIndex        =   32
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Left            =   6720
         TabIndex        =   31
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label Label8 
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
         Left            =   4080
         TabIndex        =   8
         Top             =   3000
         Width           =   750
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
         Left            =   4080
         TabIndex        =   7
         Top             =   2160
         Width           =   945
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
         Left            =   4080
         TabIndex        =   6
         Top             =   1440
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
         Left            =   4080
         TabIndex        =   5
         Top             =   720
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
         Left            =   120
         TabIndex        =   4
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
         Left            =   120
         TabIndex        =   3
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
         Left            =   120
         TabIndex        =   2
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
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   990
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "PROFESORES EXISTENTES..."
      Height          =   195
      Left            =   6240
      TabIndex        =   27
      Top             =   5040
      Width           =   2280
   End
End
Attribute VB_Name = "BASE_PROF"
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
Text4.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo2.SetFocus
End If
End Sub

Private Sub Command5_Click()
I = 0
PASSW.Show 1
If I = 1 Then
Unload Me
RETIS.Show
End If
End Sub

Private Sub Command6_Click()
'Dim profe As maestropro
If Val(Text8.Text) = 0 Then
    MsgBox "NO HAY INFORMACION PARA CONSULTAR", 64, "CONSULTAR"
    Exit Sub
End If
LIST_PRO.MATI17.Row = 0
LIST_PRO.MATI17.Col = 0
LIST_PRO.MATI17.ColWidth(0) = 500
LIST_PRO.MATI17.CellForeColor = RGB(255, 255, 255)
LIST_PRO.MATI17.CellBackColor = RGB(0, 0, 150)
LIST_PRO.MATI17.Text = "No."
LIST_PRO.MATI17.Col = 1
LIST_PRO.MATI17.ColWidth(1) = 4000
LIST_PRO.MATI17.CellForeColor = RGB(255, 255, 255)
LIST_PRO.MATI17.CellBackColor = RGB(0, 0, 150)
LIST_PRO.MATI17.Text = "NOMBRES Y APELLIDOS"
LIST_PRO.MATI17.Col = 2
LIST_PRO.MATI17.ColWidth(2) = 4200
LIST_PRO.MATI17.CellForeColor = RGB(255, 255, 255)
LIST_PRO.MATI17.CellBackColor = RGB(0, 0, 150)
LIST_PRO.MATI17.Text = "TITULO"
LIST_PRO.MATI17.Col = 3
LIST_PRO.MATI17.ColWidth(3) = 4200
LIST_PRO.MATI17.CellForeColor = RGB(255, 255, 255)
LIST_PRO.MATI17.CellBackColor = RGB(0, 0, 150)
LIST_PRO.MATI17.Text = "DIRECCION"
LIST_PRO.MATI17.Col = 4
LIST_PRO.MATI17.ColWidth(4) = 1300
LIST_PRO.MATI17.CellForeColor = RGB(255, 255, 255)
LIST_PRO.MATI17.CellBackColor = RGB(0, 0, 150)
LIST_PRO.MATI17.Text = "TELEFONO"
NAR = FreeFile
r = 0
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
While Not EOF(NAR)
    r = r + 1
    Get #NAR, r, profe
Wend
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
For I = 1 To (r - 1)
    Get #NAR, I, profe
    LIST_PRO.MATI17.Rows = I + 1
    LIST_PRO.MATI17.TextMatrix(I, 0) = I
    LIST_PRO.MATI17.TextMatrix(I, 1) = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
    LIST_PRO.MATI17.TextMatrix(I, 2) = RTrim(profe.especiali)
    LIST_PRO.MATI17.TextMatrix(I, 3) = RTrim(profe.direccion)
    LIST_PRO.MATI17.TextMatrix(I, 4) = RTrim(profe.Telefono)
Next I
Close #NAR
LIST_PRO.Show
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Agrega información de profesores a la base de datos principal."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text10_Change()
If Len(Text10.Text) = 2 Then
Text11.SetFocus
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text10.Text) = 1 Then
        Text10.Text = "0" & Text10.Text
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

Private Sub Text11_Change()
If Len(Text11.Text) = 2 Then
Text12.SetFocus
End If
If Len(Text11.Text) = 0 Then
Text10.SetFocus
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text10.SetFocus
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

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo3.SetFocus
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
Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
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
'Dim profe As maestropro
If (RTrim(Text1.Text) = "") Or (RTrim(Text2.Text) = "") Or (RTrim(Text3.Text) = "") Or (RTrim(Text4.Text) = "") Or (RTrim(Text5.Text) = "") Or (RTrim(Text10.Text) = "") Or (RTrim(Text11.Text) = "") Or (RTrim(Text12.Text) = "") Then
    MsgBox "INFORMACION INCOMPLETA", 64, "ADVERTENCIA"
    Exit Sub
End If
If (Val(Text10.Text) < 1) Or (Val(Text10.Text) > 31) Then
    MsgBox "DIA INVALIDO", 48, "FECHA DE NACIMIENTO"
    Text10.SetFocus
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

NAR = FreeFile
sir = 0
rei = 0
NAR = FreeFile
Open Ruta & "infcdpro.edu" For Random As #NAR Len = 2
While Not EOF(NAR)
sir = sir + 1
Get #NAR, sir, clat
If clat <> 0 Then
r = clat
clat = 0
Put #NAR, sir, clat
Close #NAR
rei = 1
GoTo oto
End If
Wend
Close #NAR
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
oto:

If Len(Text10.Text) = 1 Then
Text10.Text = "0" & Text10.Text
End If
If Len(Text11.Text) = 1 Then
Text11.Text = "0" & Text11.Text
End If
profe.nombres = Format(Text1.Text, ">")
profe.apellidos = Format(Text2.Text, ">")
profe.documento = Text3.Text
profe.fech_nacim = Text10.Text & "/" & Text11.Text & "/" & Text12.Text
profe.rh = Combo1.Text
profe.direccion = Text4.Text
profe.Telefono = Text5.Text
profe.año_ingre = Combo3.Text
profe.especiali = Format(Text7.Text, ">")
profe.escalafon = Combo2.Text
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Put #NAR, r, profe
Close #NAR
If rei = 0 Then
r = r + 1
Open Ruta & "contpro.edu" For Output As #NAR
Print #NAR, r
Close #NAR
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text7.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text1.SetFocus
Text8.Text = Text8.Text + 1
End Sub

Private Sub Command2_Click()
CORR_PRO.Show 1
End Sub

Private Sub Command3_Click()
'Dim profe As maestropro
If Text9.Text = "" Then
    MsgBox "ESCRIBA EL NUMERO DEL PROFESOR", 16, "ADVERTENCIA"
    Text9.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
w = Val(Text9.Text)
If ((w > r - 1) Or (w < 1)) Then
    MsgBox "REGISTRO NO EXISTE", 32, "CONSULTA DE PROFESOR"
    Text9.SetFocus
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, w, profe
Close #NAR
If ((RTrim(profe.nombres) = "") And (RTrim(profe.especiali) = "")) Then
    MsgBox "REGISTRO NO EXISTE", 16, "CONSULTAR"
    Text9.SetFocus
    Exit Sub
End If
CONS_PRO.Text1.Text = RTrim(profe.nombres)
CONS_PRO.Text2.Text = RTrim(profe.apellidos)
CONS_PRO.Text3.Text = RTrim(profe.documento)
CONS_PRO.Text11.Text = RTrim(profe.fech_nacim)
CONS_PRO.Text4.Text = RTrim(profe.rh)
CONS_PRO.Text5.Text = RTrim(profe.direccion)
CONS_PRO.Text6.Text = RTrim(profe.Telefono)
CONS_PRO.Text7.Text = RTrim(profe.año_ingre)
CONS_PRO.Text8.Text = RTrim(profe.especiali)
CONS_PRO.Text10.Text = RTrim(profe.escalafon)
CONS_PRO.Text9.Text = Text9.Text
If Dir(Ruta & "FOTOPRO\" & w & ".jpg") <> "" Then
CONS_PRO.picture2.Picture = LoadPicture(Ruta & "FOTOPRO\" & w & ".jpg")
End If
CONS_PRO.Show
End Sub

Private Sub Form_Load()
For J = 2000 To 2100
Combo3.AddItem J
Next J
'Combo3.Text = Combo3.List(0)
Combo3.Text = Combo3.List(Right(Year(Date), 3))
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
sir = 0
SIRO = 0
Open Ruta & "infcdpro.edu" For Random As #NAR Len = 2
While Not EOF(NAR)
sir = sir + 1
Get #NAR, sir, clat
If clat <> 0 Then
SIRO = SIRO + 1
End If
Wend
Close #NAR
Text8.Text = (r - 1) - SIRO
Text1.MaxLength = 20
Text2.MaxLength = 20
Text3.MaxLength = 10
Text4.MaxLength = 40
Text5.MaxLength = 12
Text7.MaxLength = 40
Text9.MaxLength = 3
Text10.MaxLength = 2
Text11.MaxLength = 2
Text12.MaxLength = 4
End Sub
