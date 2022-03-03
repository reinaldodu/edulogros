VERSION 5.00
Begin VB.Form DISCO_PROFES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación datos profesor"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "DISCO_PROFES.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Passwords"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4095
      Begin VB.TextBox Text2 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
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
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PROFESOR No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CREAR DATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2025
   End
End
Attribute VB_Name = "DISCO_PROFES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim CLAV As CLAVEPRO
'Dim profe As maestropro
'Dim argra As areagr
'Dim icur As inforcur
If Text1.Text = "" Then
    MsgBox "ESCRIBA EL NUMERO DE PROFESOR", 64, "ADVERTENCIA"
    Text1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "ESCRIBA EL PASSWORD DEL PROFESOR", 64, "ADVERTENCIA"
    Text2.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
w = Val(Text1.Text)
If ((w > r - 1) Or (w < 1)) Then
    MsgBox "PROFESOR NO EXISTE", 32, "DISCO-DATOS"
    Text1.SetFocus
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, w, profe
Close #NAR
If ((RTrim(profe.nombres) = "") And (RTrim(profe.especiali) = "")) Then
    MsgBox "REGISTRO NO EXISTE", 16, "DISCO-DATOS"
    Text1.SetFocus
    Exit Sub
End If
Frame1.Caption = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
RESP = MsgBox("DESEA CREAR EL DISCO-DATOS DE " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos) & "?", vbYesNo + vbQuestion + vbDefaultButton1, "DISCO-DATOS")
If RESP = vbYes Then
    Screen.MousePointer = 11
    'CREA EL DIRECTORIO DATOS
    MkDir RutaDir & "\DATOS"
    '*******HACER COPIA TEMPORAL EN C: **********
    MkDir "C:\DATOS"
    FileCopy Ruta & "INICIAL.EDU", "C:\DATOS\INICIAL.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "INFCUR.EDU", "C:\DATOS\INFCUR.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "PRINPRO.EDU", "C:\DATOS\PRINPRO.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "PRINALU.EDU", "C:\DATOS\PRINALU.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "CONT.EDU", "C:\DATOS\CONT.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "MATERIA.EDU", "C:\DATOS\MATERIA.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "AREAGRA.EDU", "C:\DATOS\AREAGRA.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "CONTPRO.EDU", "C:\DATOS\CONTPRO.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "CONF_DESEMP.EDU", "C:\DATOS\CONF_DESEMP.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
'    If Dir(Ruta & "RETIALU.EDU") <> "" Then
'        FileCopy Ruta & "RETIALU.EDU", "C:\DATOS\RETIALU.EDU"
'        If Err.Number <> 0 Then
'            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'    End If
    'SI EXISTE ARCHIVO DE CONFIGURACIÓN DE PORCENTAJES DE LOGROS, LO COPIA
    If Dir(Ruta & "conf_logro.edu") <> "" Then
        FileCopy Ruta & "conf_logro.edu", "C:\DATOS\conf_logro.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    FileCopy Ruta & "webhelp.txt", "C:\DATOS\webhelp.txt"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    que = 0
    Open Ruta & "CLAPRO.EDU" For Random As #NAR Len = Len(CLAV)
    While Not EOF(NAR)
        que = que + 1
        Get #NAR, que, CLAV
        If CLAV.NUMERO = w Then
            GoTo SAIRS
        End If
    Wend
SAIRS:
    ' GUARDA LA CLAVE DE ACCESO TAMBIEN EN EL SISTEMA PRINCIPAL
    CLAV.NUMERO = w
    CLAV.PASSWW = RTrim(Text2.Text)
    Put #NAR, que, CLAV
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Close #NAR
        Exit Sub
    End If
    Close #NAR
    ' CREA ARCHIVO QUE CONTIENE LA CLAVE DE ACCESO
    Open "C:\DATOS\CLAPRO.EDU" For Random As #NAR Len = Len(CLAV)
    Put #NAR, 1, CLAV
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Close #NAR
        Exit Sub
    End If
    Close #NAR
    
    'COPIA ARCHIVO DE CONTROL DE BLOQUEO DE LOGROS
    'Open "C:\DATOS\periodosL.edu" For Output As #NAR
    'Write #NAR, "1", "1", "1", "1", "1"
    'Close #NAR
    FileCopy Ruta & "periodosL.edu", "C:\DATOS\periodosL.edu"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'COPIA ARCHIVO DE CONTROL DE BLOQUEO DE DESEMPEÑOS
    'Open "C:\DATOS\periodosD.edu" For Output As #NAR
    'Write #NAR, "1", "1", "1", "1", "1"
    'Close #NAR
    FileCopy Ruta & "periodosD.edu", "C:\DATOS\periodosD.edu"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    ' COPIAR LOS GRUPOS EN DONDE DA CLASE EL PROFESOR
    CERD = 0
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        CERD = CERD + 1
        Get #NAR, CERD, argra
        If argra.num_pro = w Then
            FileCopy Ruta & RTrim(argra.nom_grup) & ".gru", "C:\DATOS\" & RTrim(argra.nom_grup) & ".gru"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
    Wend
    Close #NAR
    ' SI ES DIRECTOR DE GRUPO COPIA EL ARCHIVO DEL GRUPO
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        If icur.director = w Then
            If Dir(Ruta & RTrim(icur.nom) & ".gru") <> "" Then
                FileCopy Ruta & RTrim(icur.nom) & ".gru", "C:\DATOS\" & RTrim(icur.nom) & ".gru"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
        End If
    Wend
    Close #NAR
    'COPIAR ARCHIVOS DE LOGROS, DESEMPEÑOS, NOTAS Y PORCENTAJES DE LOGROS (SI EXISTEN)
    For lw = 1 To 4
        CERD = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            CERD = CERD + 1
            Get #NAR, CERD, argra
            If argra.num_pro = w Then
                If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".dsp") <> "" Then
                    FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".dsp", "C:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".dsp"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                End If
                If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs") <> "" Then
                    FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs", "C:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                End If
                If Dir(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> "" Then
                    If FileLen(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> 0 Then
                        FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr", "C:\DATOS\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr"
                        If Err.Number <> 0 Then
                            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                            Screen.MousePointer = 0
                            Close #NAR
                            Exit Sub
                        End If
                    End If
                End If
                
                If Dir(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ptj") <> "" Then
                    'If FileLen(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ptj") <> 0 Then
                        FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ptj", "C:\DATOS\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ptj"
                        If Err.Number <> 0 Then
                            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                            Screen.MousePointer = 0
                            Close #NAR
                            Exit Sub
                        End If
                    'End If
                End If
                
                'Copiar competencias
                If Dir(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".cpt") <> "" Then
                    If FileLen(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".cpt") <> 0 Then
                        FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".cpt", "C:\DATOS\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".cpt"
                        If Err.Number <> 0 Then
                            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                            Screen.MousePointer = 0
                            Close #NAR
                            Exit Sub
                        End If
                    End If
                End If
                
                'Copiar contenidos
                If Dir(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ctd") <> "" Then
                    If FileLen(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ctd") <> 0 Then
                        FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ctd", "C:\DATOS\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ctd"
                        If Err.Number <> 0 Then
                            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                            Screen.MousePointer = 0
                            Close #NAR
                            Exit Sub
                        End If
                    End If
                End If
                
                'Copiar ejes temáticos
                If Dir(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".eje") <> "" Then
                    If FileLen(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".eje") <> 0 Then
                        FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".eje", "C:\DATOS\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".eje"
                        If Err.Number <> 0 Then
                            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                            Screen.MousePointer = 0
                            Close #NAR
                            Exit Sub
                        End If
                    End If
                End If
                
                'Copiar planeador
                If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".pln") <> "" Then
                    FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".pln", "C:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".pln"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                End If
                
            End If
        Wend
        Close #NAR
    Next lw
    '***** COPIAR DATOS DE C:\DATOS A RUTADIR + \DATOS ******
    NGRA = Dir("C:\DATOS\*.*")
    Do While NGRA <> ""
        FileCopy "C:\DATOS\" & NGRA, RutaDir & "\DATOS\" & NGRA
        Kill "C:\DATOS\" & NGRA
        NGRA = Dir
    Loop
    RmDir ("C:\DATOS")
    MsgBox "LA INFORMACIÓN SE COPIÓ CON EXITO", 64, "Copiar Datos"
    Screen.MousePointer = 0
    Unload Me
End If
End Sub

Private Sub Command2_Click()
'Dim CLAV As CLAVEPRO
'Dim profe As maestropro
I = 0
PASSW.Show 1
If I = 1 Then
    PASSWS_PROFES.MATI14.ColWidth(0) = 4000
    PASSWS_PROFES.MATI14.ColWidth(1) = 2000
    PASSWS_PROFES.MATI14.Row = 0
    PASSWS_PROFES.MATI14.Col = 0
    PASSWS_PROFES.MATI14.CellForeColor = RGB(255, 255, 255)
    PASSWS_PROFES.MATI14.CellBackColor = RGB(0, 0, 150)
    PASSWS_PROFES.MATI14.Text = "NOMBRE DEL PROFESOR"
    PASSWS_PROFES.MATI14.Col = 1
    PASSWS_PROFES.MATI14.CellForeColor = RGB(255, 255, 255)
    PASSWS_PROFES.MATI14.CellBackColor = RGB(0, 0, 150)
    PASSWS_PROFES.MATI14.Text = "PASSWORD"
    que = 0
    NAR = FreeFile
    Open Ruta & "CLAPRO.EDU" For Random As #NAR Len = Len(CLAV)
    While Not EOF(NAR)
        que = que + 1
        Get #NAR, que, CLAV
    Wend
    Close #NAR
    For J = 1 To que - 1
        Open Ruta & "CLAPRO.EDU" For Random As #NAR Len = Len(CLAV)
        Get #NAR, J, CLAV
        Close #NAR
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, CLAV.NUMERO, profe
        Close #NAR
        PASSWS_PROFES.MATI14.Rows = J + 1
        PASSWS_PROFES.MATI14.TextMatrix(J, 0) = RTrim(profe.nombres) & " " & RTrim(profe.apellidos) & "(" & CLAV.NUMERO & ")"
        PASSWS_PROFES.MATI14.TextMatrix(J, 1) = CLAV.PASSWW
    Next J
    PASSWS_PROFES.Show 1
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Crea el diskette de datos para cada profesor."
End Sub

Private Sub Form_Load()
If Dir(Ruta & "infcur.edu") = "" Then
    Command1.Enabled = False
    Command2.Enabled = False
Else
    Command1.Enabled = True
    'Command2.Enabled = True
    Command2.Enabled = False
End If
Text1.MaxLength = 3
Text2.MaxLength = 15
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
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
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub
