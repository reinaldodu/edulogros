VERSION 5.00
Begin VB.Form RETIS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrar profesor"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2895
   Icon            =   "RETIS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Verificar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2655
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
         Height          =   360
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PROFESOR No."
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1185
      End
   End
End
Attribute VB_Name = "RETIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim profe As maestropro
If Text1.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DEL PROFESOR", 64, "BORRAR PROFESOR"
Text1.SetFocus
Exit Sub
End If
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
h = Val(Text1.Text)
If ((h > r - 1) Or (h < 1)) Then
MsgBox "REGISTRO NO EXISTE", 32, "BORRAR"
Text1.SetFocus
Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, h, profe
Close #NAR
If RTrim(profe.nombres) = "" Then
MsgBox "REGISTRO NO EXISTE", 32, "BORRAR"
Text1.SetFocus
Exit Sub
End If
CONS_PRO.Text1.Text = RTrim(profe.nombres)
CONS_PRO.Text2.Text = RTrim(profe.apellidos)
CONS_PRO.Text3.Text = RTrim(profe.documento)
CONS_PRO.Text4.Text = RTrim(profe.rh)
CONS_PRO.Text5.Text = RTrim(profe.direccion)
CONS_PRO.Text6.Text = RTrim(profe.Telefono)
CONS_PRO.Text7.Text = RTrim(profe.año_ingre)
CONS_PRO.Text8.Text = RTrim(profe.especiali)
CONS_PRO.Text9.Text = h
CONS_PRO.Text10.Text = RTrim(profe.escalafon)
CONS_PRO.Show
End Sub

Private Sub Command2_Click()
'Dim profe As maestropro
'Dim proti As pro_reti
'Dim argra As areagr
If Text1.Text = "" Then
    MsgBox "ESCRIBA EL NUMERO DEL PROFESOR", 64, "BORRAR PROFESOR"
    Text1.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
h = Val(Text1.Text)
If ((h > r - 1) Or (h < 1)) Then
    MsgBox "REGISTRO NO EXISTE", 32, "BORRAR"
    Text1.SetFocus
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, h, profe
Close #NAR
If RTrim(profe.nombres) = "" Then
    MsgBox "REGISTRO NO EXISTE", 32, "BORRAR"
    Text1.SetFocus
    Exit Sub
End If
If Dir(Ruta & "infcur.edu") <> "" Then
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        If RTrim(icur.director) = h Then
            MsgBox "PROFESOR NO SE PUEDE BORRAR, ES DIRECTOR DEL GRUPO " & RTrim(icur.nom), 32, "ADVERTENCIA"
            Text1.SetFocus
            Close #NAR
            Exit Sub
        End If
    Wend
    Close #NAR
End If
If Dir(Ruta & "AREAGRA.EDU") <> "" Then
    cona = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.num_pro) = h Then
            MsgBox "PROFESOR NO SE PUEDE BORRAR, TIENE AREAS A SU CARGO", 32, "ADVERTENCIA"
            Text1.SetFocus
            Close #NAR
            Exit Sub
        End If
    Wend
    Close #NAR
End If
'Open Ruta & "conrepro.edu" For Input As #NAR
'Input #NAR, zu
'Close #NAR
RESP = MsgBox("DESEA BORRAR ESTE PROFESOR DE LA BASE DE DATOS?", vbYesNo + vbQuestion + vbDefaultButton1, "BORRAR PROFESOR")
If RESP = vbYes Then
'    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
'    Get #NAR, h, profe
'    Close #NAR
'    Open Ruta & "retipro.edu" For Random As #NAR Len = Len(proti)
'    proti.nombres = profe.nombres
'    proti.apellidos = profe.apellidos
'    proti.documento = profe.documento
'    proti.rh = profe.rh
'    proti.direccion = profe.direccion
'    proti.Telefono = profe.Telefono
'    proti.año_ingre = profe.año_ingre
'    proti.año_retir = Combo1.Text
'    proti.especiali = profe.especiali
'    proti.escalafon = profe.escalafon
'    Put #NAR, zu, proti
'    Close #NAR
    
    profe.nombres = ""
    profe.apellidos = ""
    profe.documento = ""
    profe.fech_nacim = ""
    profe.rh = ""
    profe.direccion = ""
    profe.Telefono = ""
    profe.año_ingre = ""
    profe.especiali = ""
    profe.escalafon = ""
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    Put #NAR, h, profe
    Close #NAR
    If Dir(Ruta & "FOTOPRO\" & h & ".jpg") <> "" Then
        Kill Ruta & "FOTOPRO\" & h & ".jpg"
    End If
    'zu = zu + 1
    'Open Ruta & "conrepro.edu" For Output As #NAR
    'Print #NAR, zu
    'Close #NAR
    
    sir = 0
    Open Ruta & "infcdpro.edu" For Random As #NAR Len = 2
    While Not EOF(NAR)
        sir = sir + 1
        Get #NAR, sir, ki
        If ki = 0 Then
            ki = h
            Put #NAR, sir, ki
            Close #NAR
            Text1.SetFocus
            Exit Sub
        End If
    Wend
    Close #NAR
    Open Ruta & "infcdpro.edu" For Random As #NAR Len = 2
    ki = h
    Put #NAR, sir, ki
    Close #NAR
    Unload Me
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Borra los profesores de la base de datos principal."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
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
Private Sub Form_Load()
Text1.MaxLength = 3
End Sub
