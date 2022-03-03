VERSION 5.00
Begin VB.Form info_adicional 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información adicional"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "info_adicional.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      Picture         =   "info_adicional.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Copiar"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   6120
      Picture         =   "info_adicional.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Buscar"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      Picture         =   "info_adicional.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Enabled         =   0   'False
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
      Left            =   4920
      Picture         =   "info_adicional.frx":0B78
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Guardar"
      Top             =   480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1455
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   320
         Left            =   1560
         TabIndex        =   1
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   320
         Left            =   960
         TabIndex        =   0
         Top             =   960
         Width           =   495
      End
      Begin VB.Line Line1 
         DrawMode        =   4  'Mask Not Pen
         X1              =   0
         X2              =   2520
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Carnet No."
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   765
      End
   End
   Begin VB.TextBox informe 
      Enabled         =   0   'False
      Height          =   3375
      Left            =   120
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   7215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1545
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label grupo 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   960
      Width           =   45
   End
   Begin VB.Label nombre 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2640
      TabIndex        =   10
      Top             =   600
      Width           =   45
   End
   Begin VB.Label apellido 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "GRUPO:"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   735
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
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   990
   End
   Begin VB.Label Label1 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "info_adicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Modifico_info As Boolean

Private Sub Command1_Click()
If Modifico_info = True Then
    RESP = MsgBox("Desea guardar la información de " & RTrim(alumno.nombres) & "?", vbYesNo + vbQuestion + vbDefaultButton1, "Guardar")
    If RESP = vbYes Then
        Call Command6_Click
    Else
        Modifico_info = False
    End If
End If
If Text2.Text = "" Then
    MsgBox "Escriba un número de carnet", 32, "Advertencia"
    Text2.SetFocus
    Exit Sub
End If
If Val(Text2.Text) > 32000 Then
    MsgBox "Número de carnet inválido", 32, "Advertencia"
    Text2.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
h = Val(Text2.Text)
If ((h > I - 1) Or (h < 1)) Then
    MsgBox "Registro no existe", 32
    VERI = 0
    Text2.SetFocus
    Exit Sub
End If
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
Open Ruta & "informe.edu" For Random As #NAR Len = Len(detalle)
Get #NAR, h, detalle
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
    MsgBox "REGISTRO NO EXISTE", 32
    Text2.SetFocus
    Exit Sub
End If
Modifico_info = False
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
Get #NAR, h, aluper
Close #NAR
Command2.Enabled = True
Command4.Enabled = True
Command6.Enabled = True
informe.Enabled = True
info_adicional.Caption = "Información adicional - Carnet No." & alumno.n_carnet
apellido.Caption = RTrim(alumno.apellidos)
nombre.Caption = RTrim(alumno.nombres)
grupo.Caption = RTrim(aluper.grupo)
informe.Text = RTrim(detalle.info)
If Dir(Ruta & "FOTOALU\" & h & ".jpg") <> "" Then
    Image1.Picture = LoadPicture(Ruta & "FOTOALU\" & h & ".jpg")
Else
    Image1.Picture = LoadPicture()
End If
Text2.Text = ""
Text2.SetFocus
End Sub

Private Sub Command2_Click()
RESP = MsgBox("Desea imprimir la información de " & RTrim(alumno.nombres) & "?", vbYesNo + vbQuestion + vbDefaultButton1, "Imprimir")
If RESP = vbYes Then
    ScaleMode = 7
    Printer.ScaleMode = 7
    ttl = 1
    cant = 1
    inicia = 1
    Printer.CurrentY = 3
    Printer.CurrentX = 2.5
    Printer.Font.Size = 10
    Printer.FontBold = True
    Printer.Print apellido.Caption & " " _
    & nombre.Caption & " - (" & grupo.Caption & ")."
    Printer.FontBold = False
    Printer.Print ""
    Printer.Print ""
    While (Len(RTrim(informe.Text)) >= ttl)
        text_size = Mid(informe.Text, inicia, cant)
        letra_act = Mid(informe.Text, ttl, 1)
        If (letra_act = Chr(13)) Then
            ttl = ttl + 1
            inicia = ttl + 1
            cant = 0
            Printer.CurrentX = 2.5
            Printer.Print Trim(text_size)
        End If
        tamaño = TextWidth(text_size)
        If ((TextWidth(text_size) > 12) And (letra_act = " ")) Then
            inicia = ttl
            cant = 0
            Printer.CurrentX = 2.5
            Printer.Print Trim(text_size)
        End If
        ttl = ttl + 1
        cant = cant + 1
    Wend
    Printer.CurrentX = 2.5
    Printer.Print Trim(text_size)
    Printer.EndDoc
End If
End Sub

Private Sub Command3_Click()
BuscAlum.Show 1
End Sub

Private Sub Command4_Click()
If RTrim(informe.Text) = "" Then
    MsgBox "No existe información para copiar", 48, "Copiar"
    Exit Sub
End If
Clipboard.Clear
cop = ""
cop = apellido.Caption & " " & nombre.Caption & " - (" & grupo.Caption & ")." & vbCrLf & vbCrLf & informe.Text
Clipboard.SetText cop
End Sub

Private Sub Command6_Click()
detalle.info = informe.Text
NAR = FreeFile
Open Ruta & "informe.edu" For Random As #NAR Len = Len(detalle)
Put #NAR, h, detalle
Close #NAR
Modifico_info = False
End Sub

Private Sub Form_Load()
Modifico_info = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Modifico_info = True Then
    RESP = MsgBox("Desea guardar la información de " & RTrim(alumno.nombres) & "?", vbYesNo + vbQuestion + vbDefaultButton1, "Guardar")
    If RESP = vbYes Then
        Call Command6_Click
    End If
End If
End Sub

Private Sub informe_KeyPress(KeyAscii As Integer)
Modifico_info = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
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
