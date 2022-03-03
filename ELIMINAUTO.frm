VERSION 5.00
Begin VB.Form ELIMINAUTO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3. Retiro automático"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2430
   ControlBox      =   0   'False
   Icon            =   "ELIMINAUTO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2430
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox SigueGrado 
      Height          =   315
      ItemData        =   "ELIMINAUTO.frx":0442
      Left            =   120
      List            =   "ELIMINAUTO.frx":0470
      TabIndex        =   3
      Text            =   "PREJARDIN"
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "AÑO DE RETIRO"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "ELIMINAUTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim retiros As retiro
'Dim aluper As pertgrup
RESP = MsgBox("3. DESEA RETIRAR LOS ESTUDIANTES SIN GRUPO, PARA EL AÑO " & Text1.Text & "?", vbYesNo + vbQuestion + vbDefaultButton1, "RETIRO AUTOMATICO")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "cont.edu" For Input As #NAR
    Input #NAR, I
    Close #NAR
    Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
    For h = 1 To (I - 1)
        Get #NAR, h, aluper
        NAR = FreeFile
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, h, alumno
        Close #NAR
        NAR = NAR - 1
        'SE RETIRAN ESTUDIANTES SIN GRUPO Y QUE PERTENEZCAN AL GRADO UNDECIMO
        If (RTrim(aluper.grupo) = "SIN GRUPO") Or (RTrim(alumno.grado) = "UNDECIMO") Then
'            NAR = FreeFile
'            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'            Get #NAR, h, alumno
'            Close #NAR
'            J = 0
'            k = 0
'            NAR = FreeFile
'            Open Ruta & "retialu.edu" For Random As #NAR Len = Len(retiros)
'            While Not EOF(NAR)
'                J = J + 1
'                Get #NAR, J, retiros
'                If RTrim(retiros.nombres) = "" Then
'                    retiros.nombres = alumno.nombres
'                    retiros.apellidos = alumno.apellidos
'                    retiros.direccion = alumno.direccion
'                    retiros.Telefono = alumno.tel_acu
'                    retiros.jornada = alumno.jornada
'                    retiros.año_ingreso = alumno.año_ingre
'                    retiros.grado = alumno.grado
'                    retiros.año_retiro = Text1.Text
'                    Put #NAR, J, retiros
'                    Close #NAR
'                    k = 1
'                    GoTo sinki2
'                End If
'            Wend
'            retiros.nombres = alumno.nombres
'            retiros.apellidos = alumno.apellidos
'            retiros.direccion = alumno.direccion
'            retiros.Telefono = alumno.tel_acu
'            retiros.jornada = alumno.jornada
'            retiros.año_ingreso = alumno.año_ingre
'            retiros.grado = alumno.grado
'            retiros.año_retiro = Text1.Text
'            Put #NAR, J, retiros
'            Close #NAR
'sinki2:
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
            NAR = FreeFile
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Put #NAR, h, alumno
            Close #NAR
            NAR = NAR - 1
            If Dir(Ruta & "FOTOALU\" & h & ".jpg") <> "" Then
                Kill Ruta & "FOTOALU\" & h & ".jpg"
            End If
'            If k = 0 Then
'                Open Ruta & "contreti.edu" For Input As #NAR
'                Input #NAR, zi
'                Close #NAR
'                zi = zi + 1
'                Open Ruta & "contreti.edu" For Output As #NAR
'                Print #NAR, zi
'                Close #NAR
'            End If
'            If k = 1 Then
'                Open Ruta & "conelire.edu" For Input As #NAR
'                Input #NAR, z
'                Close #NAR
'                z = z - 1
'                Open Ruta & "conelire.edu" For Output As #NAR
'                Print #NAR, z
'                Close #NAR
'            End If
            sir = 0
            NAR = FreeFile
            Open Ruta & "infcaret.edu" For Random As #NAR Len = 2
            While Not EOF(NAR)
                sir = sir + 1
                Get #NAR, sir, clat
                If clat = 0 Then
                    clat = h
                    Put #NAR, sir, clat
                    Close #NAR
                    GoTo otro
                End If
            Wend
            clat = h
            Put #NAR, sir, clat
            Close #NAR
otro:
            NAR = NAR - 1
        End If
        aluper.grupo = "PENDIENTE"
        Put #NAR, h, aluper
    Next h
    Close #NAR
    Open Ruta & "historia\" & Text1.Text & "\" & Text1.Text & ".fin" For Random As #NAR Len = 2
    Close #NAR
    Screen.MousePointer = 0
Else
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "cont.edu" For Input As #NAR
    Input #NAR, I
    Close #NAR
    Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
    For h = 1 To (I - 1)
        aluper.grupo = "PENDIENTE"
        Put #NAR, h, aluper
    Next h
    Close #NAR
    Open Ruta & "historia\" & Text1.Text & "\" & Text1.Text & ".fin" For Random As #NAR Len = 2
    Close #NAR
    Screen.MousePointer = 0
End If
Screen.MousePointer = 11
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
'Reinicia archivo de informacion de colegios sino existe
If Dir(Ruta & "infcol.edu") = "" Then
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
End If
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
For r = 1 To (I - 1)
    Get #NAR, r, alumno
    NAR = FreeFile
    Open Ruta & "infcol.edu" For Random As #NAR Len = Len(newmatri)
    Get #NAR, r, newmatri
    'Registra colegio, grado, año y ciudad
    For w = 1 To 14
        If RTrim(newmatri.nombre(w)) = "" Then
            newmatri.nombre(w) = ini.nombre
            newmatri.grado(w) = alumno.grado
            newmatri.año(w) = Text1.Text
            newmatri.ciudad(w) = ini.ciudad
            Put #NAR, r, newmatri
            Exit For
        End If
    Next w
    Close #NAR
    NAR = NAR - 1
    'Cambia a los estudiantes al siguiente grado
    If RTrim(alumno.grado) <> "UNDECIMO" Then
        For w = 0 To 13
            If SigueGrado.List(w) = RTrim(alumno.grado) Then
                alumno.grado = SigueGrado.List(w + 1)
                Exit For
            End If
        Next w
    End If
    alumno.n_matricula = 0
    Put #NAR, r, alumno
Next r
Close #NAR
'inicia contador de matrículas a 1
Open Ruta & "conmatri.edu" For Output As #NAR
Print #NAR, 1
Close #NAR
Screen.MousePointer = 0
Unload Me
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Retira los alumnos 'Sin Grupo' de la base de datos principal."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
