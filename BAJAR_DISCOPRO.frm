VERSION 5.00
Begin VB.Form BAJAR_DISCOPRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bajar Datos"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "BAJAR_DISCOPRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   3375
         Begin VB.CheckBox Check1 
            Caption         =   "Si"
            Height          =   255
            Left            =   2640
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "¿Bajar sólo los logros del periodo?"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "BAJAR_DISCOPRO.frx":0442
         Left            =   1920
         List            =   "BAJAR_DISCOPRO.frx":0455
         TabIndex        =   2
         Text            =   "PRIMERO"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   3120
         Picture         =   "BAJAR_DISCOPRO.frx":0483
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "BAJAR_DISCOPRO.frx":08C5
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PERIODO          :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. PROFESOR:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   1230
      End
   End
End
Attribute VB_Name = "BAJAR_DISCOPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()

End Sub

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

Private Sub Command1_Click()
'Dim CLAV As CLAVEPRO
'Dim profe As maestropro
'Dim argra As areagr
'Dim ifnt As infornoti
If Text1.Text = "" Then
    MsgBox "ESCRIBA EL NUMERO DEL PROFESOR", 48, "BAJAR DATOS"
    Text1.SetFocus
    Exit Sub
End If
Screen.MousePointer = 11
If Combo1.Text = "PRIMERO" Then
    lw = 1
End If
If Combo1.Text = "SEGUNDO" Then
    lw = 2
End If
If Combo1.Text = "TERCERO" Then
    lw = 3
End If
If Combo1.Text = "CUARTO" Then
    lw = 4
End If
If Combo1.Text = "FINAL" Then
    lw = 5
End If
On Error Resume Next
Err.Clear
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
w = Val(Text1.Text)
If ((w > r - 1) Or (w < 1)) Then
    MsgBox "PROFESOR NO EXISTE", 32, "BAJAR DATOS"
    Text1.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, w, profe
Close #NAR
If ((RTrim(profe.nombres) = "") And (RTrim(profe.especiali) = "")) Then
    MsgBox "REGISTRO NO EXISTE", 16, "BAJAR DATOS"
    Text1.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
'SE VERIFICA SI LOS DATOS SELECCIONADOS CORRESPONDEN AL PROFESOR
Open RutaDir & "\CLAPRO.EDU" For Random As #NAR Len = Len(CLAV)
Get #NAR, 1, CLAV
Close #NAR
If CLAV.NUMERO <> w Then
    MsgBox "LOS DATOS SELECCIONADOS NO CORRESPONDEN AL NUMERO DE PROFESOR INGRESADO", 48, "ADVERTENCIA"
    Text1.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
Frame1.Caption = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
RESP = MsgBox("DESEA BAJAR LOS DATOS DE " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos) & "?", vbYesNo + vbQuestion + vbDefaultButton1, "BAJAR DISCO")
If RESP = vbYes Then
    CERD = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
    CERD = CERD + 1
    Get #NAR, CERD, argra
    If (argra.num_pro) = w Then
        'OBTENER EL NOMBRE DE LA MATERIA
        NAR = FreeFile
        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
        Get #NAR, argra.num_area, mate
        Close #NAR
        NAR = NAR - 1
        'COPIA OBSERVACIONES
        If Check1.Value = 0 Then
            If Dir(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS") = "" Then
                RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene observaciones para la materia " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
                If RESP = vbYes Then
                    GoTo vaconti
                Else
                    Close #NAR
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                'If FileLen("A:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS") = 0 Then
                If FileLen(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS") = 0 Then
                    RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene observaciones para la materia " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
                    If RESP = vbYes Then
                        GoTo vaconti
                    Else
                        Close #NAR
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
                'FileCopy "A:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS"
                FileCopy RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
        End If
vaconti:
        'COPIA DESEMPEÑOS
        If Check1.Value = 0 Then
            If Dir(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".DSP") = "" Then
                RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene desempeños para la materia " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
                If RESP = vbYes Then
                    GoTo vaconti2
                Else
                    Close #NAR
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                If FileLen(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".DSP") = 0 Then
                    RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene desempeños para la materia " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
                    If RESP = vbYes Then
                        GoTo vaconti2
                    Else
                        Close #NAR
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
                FileCopy RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".DSP", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".DSP"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
                'COPIA LOS ARCHIVOS DE VERIFICACION DE PLANILLAS CERRADAS
                If Dir(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".fnp") <> "" Then
                    FileCopy RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".fnp", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".fnp"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                End If
            End If
        End If
vaconti2:
        'COPIAR LOGROS
        fl = Left(argra.nom_grup, 1)
        If Check1.Value = 1 Then
            'If Dir("A:\DATOS\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR") = "" Then
            If Dir(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR") = "" Then
                RESP = MsgBox("No existen logros para el grado " & Format(RTrim(argra.grado), "<") & " de la materia " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "información incompleta")
                If RESP = vbYes Then
                    GoTo vaconti3
                Else
                    Close #NAR
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                'If FileLen("A:\DATOS\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR") = 0 Then
                If FileLen(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR") = 0 Then
                    RESP = MsgBox("No existen logros para el grado " & Format(RTrim(argra.grado), "<") & " de la materia " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Información incompleta")
                    If RESP = vbYes Then
                        GoTo vaconti3
                    Else
                        Close #NAR
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
                'FileCopy "A:\DATOS\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR"
                FileCopy RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
        End If
vaconti3:
        'COMENTARIOS EN EL INFORME FINAL.
        If Check1.Value = 0 Then
            If Dir(RutaDir & "\LRF" & RTrim(argra.nom_grup) & ".LRF") <> "" Then
                FileCopy RutaDir & "\LRF" & RTrim(argra.nom_grup) & ".LRF", Ruta & "LRF" & RTrim(argra.nom_grup) & ".LRF"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            If Dir(RutaDir & "\ORF" & RTrim(argra.nom_grup) & ".ORF") <> "" Then
                FileCopy RutaDir & "\ORF" & RTrim(argra.nom_grup) & ".ORF", Ruta & "ORF" & RTrim(argra.nom_grup) & ".ORF"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            'COPIA ARCHIVOS DE PORCENTAJES DE LOGROS SI EXISTEN
            If Dir(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".PTJ") <> "" Then
                FileCopy RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".PTJ", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".PTJ"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            
            '****** INFORMACIÓN DEL PLANEADOR ******
            'COPIA COMPENTENCIAS
            If Dir(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".CPT") <> "" Then
                FileCopy RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".CPT", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".CPT"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            'COPIA CONTENIDOS
            If Dir(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".CTD") <> "" Then
                FileCopy RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".CTD", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".CTD"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            'COPIA EJES TEMÁTICOS
            If Dir(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".EJE") <> "" Then
                FileCopy RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".EJE", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".EJE"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            'COPIA PLANEADOR
            If Dir(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".PLN") <> "" Then
                FileCopy RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".PLN", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".PLN"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            
        End If
    End If
    Wend
    Close #NAR
    'GUARDAR REGISTRO (LOG) DE ACTUALIZACION DE NOTAS EN EL SISTEMA
    ifnt.numprofe = w
    ifnt.periodo = Combo1.Text
    ifnt.fecha = Date
    ifnt.hora = Time
    Open Ruta & "infnota.edu" For Append As #NAR
    Write #NAR, ifnt.numprofe, ifnt.periodo, ifnt.fecha, ifnt.hora
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Close #NAR
        Exit Sub
    End If
    Close #NAR
    MsgBox "COPIA EXITOSA", 48, "BAJAR DATOS"
    Screen.MousePointer = 0
    'Unload Me
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Copia la información que contiene el dispositivo de datos del profesor al sistema principal."
End Sub

Private Sub Form_Load()
Text1.MaxLength = 3
'Check1.Value = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
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
