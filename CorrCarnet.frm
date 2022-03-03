VERSION 5.00
Begin VB.Form CorrCarnet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de archivo para Carnets"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "CorrCarnet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox SELCAR 
      Height          =   315
      ItemData        =   "CorrCarnet.frx":0442
      Left            =   960
      List            =   "CorrCarnet.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Generar Archivo"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ComboBox NumCar 
      Height          =   315
      ItemData        =   "CorrCarnet.frx":0467
      Left            =   1560
      List            =   "CorrCarnet.frx":04A7
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Eliminar campo"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Adicionar campo"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Campos seleccionados"
      Height          =   2415
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   2055
      Begin VB.ListBox List2 
         Height          =   2010
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nombre de campos"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      Begin VB.ListBox List1 
         Height          =   2010
         ItemData        =   "CorrCarnet.frx":04F2
         Left            =   120
         List            =   "CorrCarnet.frx":04F4
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Carnets por hoja..."
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1290
   End
End
Attribute VB_Name = "CorrCarnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
I = 0
While I < List1.ListCount
   If List1.Selected(I) = True Then
        List2.AddItem List1.List(I)
        List1.RemoveItem I
   Else
        I = I + 1
   End If
Wend
End Sub

Private Sub Command2_Click()
I = 0
While I < List2.ListCount
   If List2.Selected(I) = True Then
        List1.AddItem List2.List(I)
        List2.RemoveItem I
   Else
        I = I + 1
   End If
Wend
End Sub

Private Sub Command3_Click()
Dim AcuArchi As String, ConTab As Byte, AcuVar As String
Dim CarArch As String, CarCont As String
If List2.ListCount < 1 Then
    MsgBox "No existen campos seleccionados para generar el archivo", 64
    Exit Sub
End If
If SELCAR.Text = "ESTUDIANTES" Then
    CarArch = App.Path & "\CarnetsAlumnos.txt"
    CarCont = Ruta & "cont.edu"
End If
If SELCAR.Text = "DOCENTES" Then
    CarArch = App.Path & "\CarnetsDocentes.txt"
    CarCont = Ruta & "contpro.edu"
End If
If SELCAR.Text = "OTROS" Then
    CarArch = App.Path & "\CarnetsOtros.txt"
    CarCont = Ruta & "contotro.edu"
End If
On Error Resume Next
Err.Clear

'If Dir("c:\mis documentos\CarnetsAlumnos.txt") <> "" Then
'    Kill "c:\mis documentos\CarnetsAlumnos.txt"
If Dir(CarArch) <> "" Then
    Kill CarArch

    If Err.Number = 75 Then
        MsgBox "No se puede generar el archivo, está en uso por otra aplicación", 64
        Exit Sub
    End If
End If
Screen.MousePointer = 11
NAR = FreeFile
AcuArchi = ""
For I = 1 To Val(NumCar.Text)
    For J = 0 To List2.ListCount - 1
        If (J = List2.ListCount - 1) And (I = Val(NumCar.Text)) Then
            AcuArchi = AcuArchi & List2.List(J) & I
        Else
            AcuArchi = AcuArchi & List2.List(J) & I & Chr$(9)
        End If
    Next J
Next I
'Open "c:\mis documentos\CarnetsAlumnos.txt" For Append As #NAR
Open CarArch For Append As #NAR
If Err.Number = 70 Then
    Screen.MousePointer = 0
    MsgBox "No se puede generar el archivo, está en uso por otra aplicación", 64
    Exit Sub
End If
Print #NAR, AcuArchi
Close #NAR
ConTab = 1
AcuArchi = ""
'Open ruta & "cont.edu" For Input As #NAR
Open CarCont For Input As #NAR
Input #NAR, k
Close #NAR
If SELCAR.Text = "ESTUDIANTES" Then
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
            NAR = FreeFile
            curcar = 0
            Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
            While Not EOF(NAR)
                curcar = curcar + 1
                Get #NAR, curcar, alugru
            Wend
            Close #NAR
            Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
            For h = 1 To (curcar - 1)
                Get #NAR, h, alugru
                NAR = FreeFile
                Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
                Get #NAR, Val(alugru.num_carnet), alumno

                If (RTrim(alumno.nombres) <> "") And (RTrim(alumno.apellidos) <> "") And (RTrim(alumno.grado) <> "SIN GRADO") Then
                For J = 0 To List2.ListCount - 1
                    If List2.List(J) = "Nombres" Then _
                        AcuVar = RTrim(alumno.nombres)
                    If List2.List(J) = "Apellidos" Then _
                        AcuVar = RTrim(alumno.apellidos)
                    If List2.List(J) = "Carnet" Then _
                        AcuVar = RTrim(alumno.n_carnet)
                    If List2.List(J) = "Documento" Then _
                        AcuVar = RTrim(alumno.documento)
                    If List2.List(J) = "R.H." Then _
                        AcuVar = RTrim(alumno.rh)
                    If List2.List(J) = "Jornada" Then _
                        AcuVar = RTrim(alumno.jornada)
                    If List2.List(J) = "Grado" Then _
                        AcuVar = RTrim(alumno.grado)
                    If List2.List(J) = "F_nacimiento" Then _
                        AcuVar = RTrim(alumno.f_nacimiento)
                    If List2.List(J) = "Teléfono" Then
                        NAR = FreeFile
                        Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
                        Get #NAR, Val(alugru.num_carnet), AdiCampo
                        AcuVar = RTrim(AdiCampo.Tel_casa)
                        Close #NAR
                        NAR = NAR - 1
                    End If
                    If List2.List(J) = "Grupo" Then
                        NAR = FreeFile
                        Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
                        Get #NAR, Val(alumno.n_carnet), aluper
                        AcuVar = RTrim(aluper.grupo)
                        AcuVar = Right(AcuVar, Len(AcuVar) - 1)
                        Close #NAR
                        NAR = NAR - 1
                    End If
                    If (ConTab < Val(NumCar.Text)) Or (J < List2.ListCount - 1) Then
                        AcuArchi = AcuArchi & AcuVar & Chr$(9)
                    Else
                        AcuArchi = AcuArchi & AcuVar
                        NAR = FreeFile
                        Open CarArch For Append As #NAR
                        Print #NAR, AcuArchi
                        Close #NAR
                        NAR = NAR - 1
                        ConTab = 0
                        AcuArchi = ""
                    End If
                Next J
                ConTab = ConTab + 1
            End If
            Close #NAR
            NAR = NAR - 1
         
            Next h
            Close #NAR
            
    NAR = NAR - 1
    Wend
    Close #NAR - 1
'    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'    For h = 1 To k - 1
'        Get #NAR, h, alumno
'        If (RTrim(alumno.nombres) <> "") And (RTrim(alumno.apellidos) <> "") And (RTrim(alumno.grado) <> "SIN GRADO") Then
'                For J = 0 To List2.ListCount - 1
'                    If List2.List(J) = "Nombres" Then _
'                        AcuVar = RTrim(alumno.nombres)
'                    If List2.List(J) = "Apellidos" Then _
'                        AcuVar = RTrim(alumno.apellidos)
'                    If List2.List(J) = "Carnet" Then _
'                        AcuVar = RTrim(alumno.n_carnet)
'                    If List2.List(J) = "Documento" Then _
'                        AcuVar = RTrim(alumno.documento)
'                    If List2.List(J) = "R.H." Then _
'                        AcuVar = RTrim(alumno.rh)
'                    If List2.List(J) = "Jornada" Then _
'                        AcuVar = RTrim(alumno.jornada)
'                    If List2.List(J) = "Grado" Then _
'                        AcuVar = RTrim(alumno.grado)
'                    If List2.List(J) = "F_nacimiento" Then _
'                        AcuVar = RTrim(alumno.f_nacimiento)
'                    If List2.List(J) = "Teléfono" Then
'                        NAR = FreeFile
'                        Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
'                        Get #NAR, h, AdiCampo
'                        AcuVar = RTrim(AdiCampo.Tel_casa)
'                        Close #NAR
'                        NAR = NAR - 1
'                    End If
'                    If List2.List(J) = "Grupo" Then
'                        NAR = FreeFile
'                        Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
'                        Get #NAR, Val(Right(alumno.n_carnet, 5)), aluper
'                        AcuVar = RTrim(aluper.grupo)
'                        Close #NAR
'                        NAR = NAR - 1
'                    End If
'                    If (ConTab < Val(NumCar.Text)) Or (J < List2.ListCount - 1) Then
'                        AcuArchi = AcuArchi & AcuVar & Chr$(9)
'                    Else
'                        AcuArchi = AcuArchi & AcuVar
'                        NAR = FreeFile
'                        Open CarArch For Append As #NAR
'                        Print #NAR, AcuArchi
'                        Close #NAR
'                        NAR = NAR - 1
'                        ConTab = 0
'                        AcuArchi = ""
'                    End If
'                Next J
'                ConTab = ConTab + 1
'        End If
'    Next h
'    Close #NAR
End If

If SELCAR.Text = "DOCENTES" Then
    NAR = FreeFile
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    For h = 1 To k - 1
        Get #NAR, h, profe
        If (RTrim(profe.nombres) <> "") And (RTrim(profe.apellidos) <> "") Then
                For J = 0 To List2.ListCount - 1
                    If List2.List(J) = "Nombres" Then _
                        AcuVar = RTrim(profe.nombres)
                    If List2.List(J) = "Apellidos" Then _
                        AcuVar = RTrim(profe.apellidos)
                    If List2.List(J) = "Documento" Then _
                        AcuVar = RTrim(profe.documento)
                    If List2.List(J) = "R.H." Then _
                        AcuVar = RTrim(profe.rh)
                    If List2.List(J) = "Teléfono" Then _
                        AcuVar = RTrim(profe.Telefono)
                    If List2.List(J) = "Especialidad" Then _
                        AcuVar = RTrim(profe.especiali)
                    If (ConTab < Val(NumCar.Text)) Or (J < List2.ListCount - 1) Then
                        AcuArchi = AcuArchi & AcuVar & Chr$(9)
                    Else
                        AcuArchi = AcuArchi & AcuVar
                        NAR = FreeFile
                        Open CarArch For Append As #NAR
                        Print #NAR, AcuArchi
                        Close #NAR
                        NAR = NAR - 1
                        ConTab = 0
                        AcuArchi = ""
                    End If
                Next J
                ConTab = ConTab + 1
        End If
    Next h
    Close #NAR
End If

' OJO QUEDA PENDIENTE LA BASE DE DATOS PRINOTRO.EDU

If ConTab <> 1 Then
    For J = 1 To ((Val(NumCar.Text) * List2.ListCount) - ((ConTab - 1) * List2.ListCount) - 1)
        AcuArchi = AcuArchi & Chr$(9)
    Next J
    NAR = FreeFile
    Open CarArch For Append As #NAR
    Print #NAR, AcuArchi
    Close #NAR
End If
Screen.MousePointer = 0
MsgBox "Archivo generado con éxito", 64
Unload Me
End Sub

Private Sub Form_Load()
NumCar.Text = NumCar.List(0)
SELCAR.Text = SELCAR.List(0)
End Sub

Private Sub SELCAR_Click()
List1.Clear
If SELCAR.Text = "ESTUDIANTES" Then
    List1.AddItem "Nombres"
    List1.AddItem "Apellidos"
    List1.AddItem "Carnet"
    List1.AddItem "Documento"
    List1.AddItem "R.H."
    List1.AddItem "Jornada"
    List1.AddItem "Grado"
    List1.AddItem "Grupo"
    List1.AddItem "Teléfono"
    List1.AddItem "F_nacimiento"
    Command3.ToolTipText = "Crea el archivo " & App.Path & "\CarnetsAlumnos.txt para combinar correspondencia"
End If
If SELCAR.Text = "DOCENTES" Then
    List1.AddItem "Nombres"
    List1.AddItem "Apellidos"
    List1.AddItem "Documento"
    List1.AddItem "R.H."
    List1.AddItem "Teléfono"
    List1.AddItem "Especialidad"
    Command3.ToolTipText = "Crea el archivo " & App.Path & "\CarnetsDocentes.txt para combinar correspondencia"
End If
If SELCAR.Text = "OTROS" Then
    List1.AddItem "Nombres"
    List1.AddItem "Apellidos"
    List1.AddItem "Documento"
    List1.AddItem "R.H."
    List1.AddItem "Teléfono"
    List1.AddItem "Cargo"
    Command3.ToolTipText = "Crea el archivo " & App.Path & "\CarnetsOtros.txt para combinar correspondencia"
End If
End Sub
