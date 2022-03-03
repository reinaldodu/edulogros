VERSION 5.00
Begin VB.Form RETI_ALUM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrar Estudiante"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3045
   Icon            =   "RETI_ALUM.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Verificar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2775
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
         ForeColor       =   &H0000FFFF&
         Height          =   320
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1125
      End
   End
End
Attribute VB_Name = "RETI_ALUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim aluper As pertgrup
If Text1.Text = "" Then
MsgBox "ESCRIBA UN NUMERO DE CARNET", 64, "ADVERTENCIA"
Text1.SetFocus
Exit Sub
End If
If Val(Text1.Text) > 32000 Then
MsgBox "No. DE CARNET INVALIDO", 64, "ADVERTENCIA"
Text1.SetFocus
Exit Sub
End If
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
h = Val(Text1.Text)
If ((h > I - 1) Or (h < 1)) Then
MsgBox "REGISTRO NO EXISTE", 32, "BORRAR"
Text1.SetFocus
Exit Sub
End If
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
MsgBox "REGISTRO NO EXISTE", 32, "BORRAR"
Text1.SetFocus
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
CONS_ALUM.Show
End Sub

Private Sub Command2_Click()
'Dim alumno As maestroalum
'Dim retiros As retiro
'Dim aluper As pertgrup
If Text1.Text = "" Then
MsgBox "ESCRIBA UN NUMERO DE CARNET", 64, "ADVERTENCIA"
Text1.SetFocus
Exit Sub
End If
If Val(Text1.Text) > 32000 Then
MsgBox "No. DE CARNET INVALIDO", 64, "ADVERTENCIA"
Text1.SetFocus
Exit Sub
End If
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
h = Val(Text1.Text)
If ((h > I - 1) Or (h < 1)) Then
MsgBox "REGISTRO NO EXISTE", 32, "BORRAR"
Text1.SetFocus
Exit Sub
End If
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
MsgBox "REGISTRO NO EXISTE", 32, "BORRAR"
Text1.SetFocus
Exit Sub
End If
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
Get #NAR, h, aluper
Close #NAR
If RTrim(aluper.grupo) <> "PENDIENTE" Then
MsgBox "ESTUDIANTE NO SE PUEDE BORRAR, PERTENECE AL GRUPO " & RTrim(aluper.grupo), 32
Text1.SetFocus
Exit Sub
End If
RESP = MsgBox("DESEA BORRAR ESTE ESTUDIANTE DE LA BASE DE DATOS?", vbYesNo + vbQuestion + vbDefaultButton1, "BORRAR ESTUDIANTE")
If RESP = vbYes Then
    'Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    'Get #NAR, h, alumno
    'Close #NAR
    
    'J = 0
    'k = 0
    'Open Ruta & "retialu.edu" For Random As #NAR Len = Len(retiros)
    'While Not EOF(NAR)
    'J = J + 1
    'Get #NAR, J, retiros
    'If RTrim(retiros.nombres) = "" Then
    'retiros.nombres = alumno.nombres
    'retiros.apellidos = alumno.apellidos
    'retiros.direccion = alumno.direccion
    'retiros.Telefono = alumno.tel_acu
    'retiros.jornada = alumno.jornada
    'retiros.año_ingreso = alumno.año_ingre
    'retiros.grado = alumno.grado
    'retiros.año_retiro = Combo1.Text
    'Put #NAR, J, retiros
    'Close #NAR
    'k = 1
    'GoTo sinki
    'End If
    'Wend
    'retiros.nombres = alumno.nombres
    'retiros.apellidos = alumno.apellidos
    'retiros.direccion = alumno.direccion
    'retiros.Telefono = alumno.tel_acu
    'retiros.jornada = alumno.jornada
    'retiros.año_ingreso = alumno.año_ingre
    'retiros.grado = alumno.grado
    'retiros.año_retiro = Combo1.Text
    'Put #NAR, J, retiros
    'Close #NAR
    'sinki:
    alumno.n_carnet = ""
    alumno.n_matricula = 0
    alumno.nombres = ""
    alumno.apellidos = ""
    alumno.documento = ""
    alumno.f_nacimiento = ""
    alumno.rh = ""
    alumno.acudiente = ""
    alumno.tel_acu = ""
    alumno.padre = ""
    alumno.tel_pa = ""
    alumno.madre = ""
    alumno.tel_ma = ""
    alumno.direccion = ""
    alumno.jornada = ""
    alumno.año_ingre = ""
    alumno.grado = ""
    alumno.sexo = ""
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Put #NAR, h, alumno
    Close #NAR
    If Dir(Ruta & "FOTOALU\" & h & ".jpg") <> "" Then
    Kill Ruta & "FOTOALU\" & h & ".jpg"
    End If
    
    detalle.info = ""
    'NAR = FreeFile
    Open Ruta & "informe.edu" For Random As #NAR Len = Len(detalle)
    Put #NAR, h, detalle
    Close #NAR
    
    'If k = 0 Then
    'Open Ruta & "contreti.edu" For Input As #NAR
    'Input #NAR, zi
    'Close #NAR
    'zi = zi + 1
    'Open Ruta & "contreti.edu" For Output As #NAR
    'Print #NAR, zi
    'Close #NAR
    'End If
    'If k = 1 Then
    'Open Ruta & "conelire.edu" For Input As #NAR
    'Input #NAR, z
    'Close #NAR
    'z = z - 1
    'Open Ruta & "conelire.edu" For Output As #NAR
    'Print #NAR, z
    'Close #NAR
    'End If
    sir = 0
    Open Ruta & "infcaret.edu" For Random As #NAR Len = 2
    While Not EOF(NAR)
        sir = sir + 1
        Get #NAR, sir, clat
        If clat = 0 Then
            clat = h
            Put #NAR, sir, clat
            Close #NAR
            Text1.SetFocus
            Exit Sub
        End If
    Wend
    clat = h
    Put #NAR, sir, clat
    Close #NAR
    Unload Me
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Borra los estudiantes de la base de datos principal."
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
'For I = 1998 To 2100
'Combo1.AddItem I
'Next I
'Combo1.Text = Combo1.List(0)
Text1.MaxLength = 5
End Sub
