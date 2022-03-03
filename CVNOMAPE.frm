VERSION 5.00
Begin VB.Form CVNOMAPE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por nombres y apellidos"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "CVNOMAPE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "APELLIDOS:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRES:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "CVNOMAPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim aluper As pertgrup
If RTrim(Text1.Text) = "" Then
    MsgBox "ESCRIBA LOS NOMBRES", 64, "ADVERTENCIA"
    Text1.SetFocus
    Exit Sub
End If
If RTrim(Text2.Text) = "" Then
    MsgBox "ESCRIBA LOS APELLIDOS", 64, "ADVERTENCIA"
    Text2.SetFocus
    Exit Sub
End If
Text1.Text = RTrim(Format(Text1.Text, ">"))
Text2.Text = RTrim(Format(Text2.Text, ">"))
NAR = FreeFile
cona = 0
Screen.MousePointer = 11
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, alumno
    If (RTrim(alumno.nombres) = Text1.Text) And (RTrim(alumno.apellidos) = Text2.Text) Then
        NAR = FreeFile
        Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
        Get #NAR, Val(alumno.n_carnet), aluper
        Close #NAR
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
        If Dir(Ruta & "FOTOALU\" & Val(alumno.n_carnet) & ".jpg") <> "" Then
            CONS_ALUM.Picture1.Picture = LoadPicture(Ruta & "FOTOALU\" & Val(alumno.n_carnet) & ".jpg")
        End If
        NAR = NAR - 1
        Close #NAR
        Screen.MousePointer = 0
        Unload Me
        CONS_ALUM.Show
        Exit Sub
    End If
Wend
Close #NAR
Screen.MousePointer = 0
MsgBox "NO SE ENCONTRO NINGUN REGISTRO", 64
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta de alumnos por nombres y apellidos."
End Sub

Private Sub Form_Load()
Text1.MaxLength = 20
Text2.MaxLength = 20
If Dir(Ruta & "prinalu.edu") = "" Then
    Command1.Enabled = False
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub
