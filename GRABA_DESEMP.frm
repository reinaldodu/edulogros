VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GRABA_DESEMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grabar desempeños"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9765
   Icon            =   "GRABA_DESEMP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Reaprendizaje"
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   5160
      TabIndex        =   22
      Top             =   6000
      Width           =   2175
      Begin VB.CheckBox Check1 
         Caption         =   "Activar reaprendizaje"
         Height          =   495
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "Al activar esta casilla puede marcar/desmarcar logros con reaprendizaje dando clic en el porcentaje del logro."
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de Lista"
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   7560
      TabIndex        =   15
      Top             =   6000
      Width           =   2055
      Begin VB.CommandButton Command10 
         Caption         =   "Borrar"
         Height          =   735
         Left            =   1080
         Picture         =   "GRABA_DESEMP.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Borrar un estudiante de la lista"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Actualizar"
         Height          =   735
         Left            =   120
         Picture         =   "GRABA_DESEMP.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Actualizar la lista de estudiantes"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   6840
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   6840
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00800000&
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9495
      Begin VB.CommandButton Command8 
         Caption         =   "Pegar columna"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Pegar todo"
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   8520
         Picture         =   "GRABA_DESEMP.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir la lista de desempeños"
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Guardar"
         Height          =   615
         Left            =   7440
         Picture         =   "GRABA_DESEMP.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Guardar los desempeños"
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Ver logros"
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Pegar fila"
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Copiar"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4680
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid Mt_desemp 
         Height          =   4215
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7435
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   3
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "GRABA_DESEMP.frx":084A
      Left            =   7920
      List            =   "GRABA_DESEMP.frx":085D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   6000
      Width           =   4815
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MATERIA:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   630
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3720
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO:"
      Height          =   195
      Left            =   6960
      TabIndex        =   0
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "GRABA_DESEMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim simucopy(100) As String, kcpy As Integer, vcpy As Integer, VerInfo As Boolean, codlogro(10) As Byte

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo2.SetFocus
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo3.SetFocus
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If Command4.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command4_Click
End If
End Sub

Private Sub Command1_Click()
Dim colcopy As String
colcopy = InputBox("Código de estudiante a copiar?" & Chr(13) & "(escriba un número de 1 a " & Mt_desemp.Rows - 1 & ")", "copiar")
If colcopy = "" Then Exit Sub
If (Val(colcopy) < 1) Or (Val(colcopy) > Val(Mt_desemp.Rows - 1)) Then
    MsgBox "Código no existe", 48, "copiar"
    Exit Sub
End If
For kcpy = 3 To Val(Mt_desemp.Cols - 1)
    simucopy(kcpy) = Mt_desemp.TextMatrix(Val(colcopy), kcpy)
Next kcpy
Command2.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
End Sub

Private Sub Command10_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    If Mt_desemp.Rows = 2 Then
        MsgBox "No se puede eliminar el último alumno de la lista", 32, "Eliminar"
        Exit Sub
    End If
    TTT = InputBox("Escriba el código que desea eliminar" & Chr(13) & "(Escriba un número entre 1 y " & Mt_desemp.Rows - 1 & ")", "Eliminar alumno")
    If TTT = "" Then
        MsgBox "No escribió el código", 64, "Eliminar"
        Exit Sub
    End If
    If Val(TTT) > Val(Mt_desemp.Rows - 1) Or (Val(TTT) < 1) Then
        MsgBox "No existe este código en la lista", 32, "Eliminar"
        Exit Sub
    End If
    Mt_desemp.RemoveItem Val(TTT)
    For TT = 1 To Val(Mt_desemp.Rows - 1)
        Mt_desemp.TextMatrix(TT, 0) = TT
    Next TT
    VALI4 = False
End If
End Sub

Private Sub Command2_Click()
For kcpy = 3 To Val(Mt_desemp.Cols - 1)
        Mt_desemp.TextMatrix(Mt_desemp.Row, kcpy) = simucopy(kcpy)
Next kcpy
VALI4 = False
End Sub

Private Sub Command3_Click()
SWobserv = False
Ver_Obser.Show
End Sub

Private Sub Command4_Click()
Dim ConfLgr As Byte, VeriManual As Boolean
If VALI4 = False Then
    Call Command5_Click
End If
Unload Ver_Obser
Label6.Caption = ""
Label4.Caption = ""
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Check1.Value = 0
Check1.Enabled = False
Mt_desemp.Rows = 1
Mt_desemp.Cols = 3
If Dir(Ruta & Combo2.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
NAR = FreeFile
TN = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    TN = TN + 1
    Get #NAR, TN, mate
    If RTrim(mate.nom) = Combo3.Text Then
        que = mate.num
    End If
Wend
Close #NAR
ret = 0
Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = RTrim(Combo2.Text) Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
    End If
Wend
Close #NAR

pio = 0
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If RTrim(argra.grado) = RE22 And RTrim(argra.nom_grup) = Combo2.Text And argra.num_area = que Then
        NAR = FreeFile
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, (argra.num_pro), profe
        Close #NAR
        PRO = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        pio = 1
        NAR = NAR - 1
    End If
Wend
Close #NAR
If pio = 0 Then
    MsgBox "NO SE HA CREADO EL AREA " & Combo3.Text & " PARA ESTE GRUPO", 64, "ADVERTENCIA"
    Combo3.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If

If RTrim(Combo1.Text) = "PRIMERO" Then
    lw = 1
End If
If RTrim(Combo1.Text) = "SEGUNDO" Then
    lw = 2
End If
If RTrim(Combo1.Text) = "TERCERO" Then
    lw = 3
End If
If RTrim(Combo1.Text) = "CUARTO" Then
    lw = 4
End If
If RTrim(Combo1.Text) = "FINAL" Then
    lw = 5
End If

If JOJI = "UNICA" Then
fl = "1"
End If
If JOJI = "MAÑANA" Then
fl = "2"
End If
If JOJI = "TARDE" Then
fl = "3"
End If
If JOJI = "NOCHE" Then
fl = "4"
End If
Label4.Caption = Combo2.Text
Label6.Caption = Combo2.Text & que & lw
ser = Left(RE22, 3)
FERT = 0
Cont_Lgr = 0

If (lw <> 5) Then
    ' NO DEJA GRABAR DESEMPEÑOS SI NO SE HAN GRABADO PORCENTAJES DE LOGROS EN EL CASO QUE SE ENCUENTRE DISPONIBLE ESTA CONFIGURACION
    VeriManual = False
    If Dir(Ruta & "conf_logro.edu") <> "" Then
        Open Ruta & "conf_logro.edu" For Input As #NAR
        Input #NAR, ConfLgr
        Close #NAR
        If ConfLgr = 1 Then
            VeriManual = True
            If Dir(Ruta & fl & ser & que & lw & ".ptj") = "" Then
                MsgBox "El sistema está configurado para agregar porcentajes por logros manualmente. Debe primero ingresarlos", 64, "ADVERTENCIA"
                Combo3.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    End If
    
    Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
    While Not EOF(NAR)
        FERT = FERT + 1
        Get #NAR, FERT, logru
        If Trim(logru.indicador) = "L" Then
            Cont_Lgr = Cont_Lgr + 1
            codlogro(Cont_Lgr) = FERT
        End If
    Wend
    Close #NAR
    If Cont_Lgr = 0 Then
        MsgBox "DEBE GRABAR PRIMERO LOGROS DE " & Combo3.Text & " PARA " & Combo2.Text, 64, "ADVERTENCIA"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If Cont_Lgr > 10 Then
        MsgBox "NO SE PUEDE CALIFICAR MÁS DE 10 LOGROS POR PERIODO, VERIFIQUE LA CANTIDAD DE LOGROS CREADOS PARA " & Combo3.Text, 64, "ADVERTENCIA"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command1.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command9.Enabled = True
    Command10.Enabled = True
    Check1.Enabled = True
    'MUESTRA PORCENTAJES DE EVALUACION DE LOGROS, SI EL  SISTEMA ESTA CONFIGURADO EN PORCENTAJES MANUALES
    If VeriManual = False Then
        For h = 1 To Cont_Lgr
            Mt_desemp.Cols = Mt_desemp.Cols + 1
            Mt_desemp.ColWidth(Mt_desemp.Cols - 1) = 800
            Mt_desemp.TextMatrix(0, Mt_desemp.Cols - 1) = "Lgr No." & h
            Mt_desemp.Row = 0
            Mt_desemp.Col = (Mt_desemp.Cols - 1)
            Mt_desemp.CellForeColor = RGB(255, 255, 255)
            Mt_desemp.CellBackColor = RGB(0, 0, 150)
        Next h
    Else
        NAR = FreeFile
        Open Ruta & fl & ser & que & lw & ".ptj" For Random As #NAR Len = Len(porcent_manual)
        For h = 1 To Cont_Lgr
            Get #NAR, h, porcent_manual
            Mt_desemp.Cols = Mt_desemp.Cols + 1
            Mt_desemp.ColWidth(Mt_desemp.Cols - 1) = 800
            Mt_desemp.TextMatrix(0, Mt_desemp.Cols - 1) = "L" & h & " [" & porcent_manual.porcent_logro & "%]"
            Mt_desemp.Row = 0
            Mt_desemp.Col = (Mt_desemp.Cols - 1)
            Mt_desemp.CellForeColor = RGB(255, 255, 255)
            Mt_desemp.CellBackColor = RGB(0, 0, 150)
        Next h
        Close #NAR
    End If
    'Label4.Caption = Combo2.Text
    'Label6.Caption = Combo2.Text & que & lw
    If Dir(Ruta & Combo2.Text & que & lw & ".dsp") = "" Then
        For I = 1 To (ret - 1)
            Mt_desemp.Rows = I + 1
            Mt_desemp.TextMatrix(I, 0) = I
            Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
            Get #NAR, I, alugru
            Close #NAR
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, (Val(alugru.num_carnet)), alumno
            Close #NAR
            Mt_desemp.TextMatrix(I, 1) = alumno.n_carnet
            Mt_desemp.TextMatrix(I, 2) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        Next I
    Else
        'Command11.Enabled = True
        Y = 0
        Open Ruta & Combo2.Text & que & lw & ".dsp" For Random As #NAR Len = Len(notas_desemp)
        While Not EOF(NAR)
            Y = Y + 1
            Get #NAR, Y, notas_desemp
        Wend
        Close #NAR
        For I = 1 To (Y - 1)
            Mt_desemp.Rows = I + 1
            Mt_desemp.TextMatrix(I, 0) = I
            Open Ruta & Combo2.Text & que & lw & ".dsp" For Random As #NAR Len = Len(notas_desemp)
            Get #NAR, I, notas_desemp
            Close #NAR
            For J = 1 To (Cont_Lgr)
                If notas_desemp.porcentaje(J) = 0 Then
                    Mt_desemp.TextMatrix(I, J + 2) = ""
                Else
                    Mt_desemp.TextMatrix(I, J + 2) = notas_desemp.porcentaje(J)
                End If
                Mt_desemp.Row = I
                Mt_desemp.Col = J + 2
                If notas_desemp.recuperado(J) = True Then
                    Mt_desemp.CellFontBold = True
                    Mt_desemp.CellForeColor = RGB(255, 0, 0)
                Else
                    Mt_desemp.CellFontBold = False
                    Mt_desemp.CellForeColor = RGB(0, 0, 0)
                End If
                
            Next J
            If RTrim(notas_desemp.num_carnet) = "" Then
                GoTo salbla
            End If
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, (Val(notas_desemp.num_carnet)), alumno
            Close #NAR
            Mt_desemp.TextMatrix(I, 2) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
            Mt_desemp.TextMatrix(I, 1) = alumno.n_carnet
salbla:
        Next I
    End If
    
    'Label1.Caption = "GRABACION DE BOLETIN JORNADA:" & JOJI & "  GRADO: " & RE22
    Frame1.Caption = "GRUPO: " & Combo2.Text & " - " & " AREA: " & Combo3.Text & " - " & " PROFESOR(A): " & PRO
    'Frame2.Caption = "PERIODO " & Combo1.Text
    'Text7.Text = I - 1
    Mt_desemp.Row = 1
    Mt_desemp.Col = 1
    Mt_desemp.SetFocus
    Screen.MousePointer = 0

Else
'****** NOTAS DE NIVELACIONES  **********
    Command1.Enabled = True
    'Command3.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command9.Enabled = True
    Command10.Enabled = True
    'Check1.Enabled = True
    Mt_desemp.Cols = Mt_desemp.Cols + 2
    Mt_desemp.ColWidth(Mt_desemp.Cols - 2) = 800
    Mt_desemp.ColWidth(Mt_desemp.Cols - 1) = 800
    Mt_desemp.TextMatrix(0, Mt_desemp.Cols - 2) = "Niv.#1"
    Mt_desemp.TextMatrix(0, Mt_desemp.Cols - 1) = "Niv.#2"
    Mt_desemp.Row = 0
    Mt_desemp.Col = (Mt_desemp.Cols - 2)
    Mt_desemp.CellForeColor = RGB(255, 255, 255)
    Mt_desemp.CellBackColor = RGB(0, 0, 150)
    Mt_desemp.Col = (Mt_desemp.Cols - 1)
    Mt_desemp.CellForeColor = RGB(255, 255, 255)
    Mt_desemp.CellBackColor = RGB(0, 0, 150)
    If Dir(Ruta & Combo2.Text & que & lw & ".dsp") = "" Then
        For I = 1 To (ret - 1)
            Mt_desemp.Rows = I + 1
            Mt_desemp.TextMatrix(I, 0) = I
            Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
            Get #NAR, I, alugru
            Close #NAR
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, (Val(alugru.num_carnet)), alumno
            Close #NAR
            Mt_desemp.TextMatrix(I, 1) = alumno.n_carnet
            Mt_desemp.TextMatrix(I, 2) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        Next I
        
    Else
        'Command11.Enabled = True
        Y = 0
        Open Ruta & Combo2.Text & que & lw & ".dsp" For Random As #NAR Len = Len(notas_desemp)
        While Not EOF(NAR)
            Y = Y + 1
            Get #NAR, Y, notas_desemp
        Wend
        Close #NAR
        For I = 1 To (Y - 1)
            Mt_desemp.Rows = I + 1
            Mt_desemp.TextMatrix(I, 0) = I
            Open Ruta & Combo2.Text & que & lw & ".dsp" For Random As #NAR Len = Len(notas_desemp)
            Get #NAR, I, notas_desemp
            Close #NAR
            For J = 1 To 2
                If notas_desemp.porcentaje(J) = 0 Then
                    Mt_desemp.TextMatrix(I, J + 2) = ""
                Else
                    Mt_desemp.TextMatrix(I, J + 2) = notas_desemp.porcentaje(J)
                End If
            Next J
            If RTrim(notas_desemp.num_carnet) = "" Then
                GoTo salbla11
            End If
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, (Val(notas_desemp.num_carnet)), alumno
            Close #NAR
            Mt_desemp.TextMatrix(I, 2) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
            Mt_desemp.TextMatrix(I, 1) = alumno.n_carnet
salbla11:
        Next I
    End If
    Screen.MousePointer = 0
End If
End Sub

Private Sub Command5_Click()
'If Val(Mt_desemp.Rows - 1) = 0 Then
'    MsgBox "ESCOJA EL NOMBRE DEL GRUPO, EL AREA Y PRESIONE OK", 48, "GUARDAR"
'    Combo2.SetFocus
'    Exit Sub
'End If
RESP = MsgBox("DESEA GUARDAR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
    Screen.MousePointer = 11
    If Dir(Ruta & Label6.Caption & ".dsp") <> "" Then
        Kill Ruta & Label6.Caption & ".dsp"
    End If
    VerInfo = False
    For I = 1 To (Mt_desemp.Rows - 1)
        For J = 3 To Mt_desemp.Cols - 1
            If Mt_desemp.TextMatrix(I, J) <> "" Then
                VerInfo = True
            End If
        Next J
    Next I
    If VerInfo = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    NAR = FreeFile
    Open Ruta & Label6.Caption & ".dsp" For Random As #NAR Len = Len(notas_desemp)
    For I = 1 To (Mt_desemp.Rows - 1)
        For J = 3 To Mt_desemp.Cols - 1
            If Trim(Mt_desemp.TextMatrix(I, J)) = "" Then
                notas_desemp.porcentaje(J - 2) = 0
            Else
                notas_desemp.porcentaje(J - 2) = Mt_desemp.TextMatrix(I, J)
            End If
            'Verificar los logros recuperados
            Mt_desemp.Row = I
            Mt_desemp.Col = J
            If Mt_desemp.CellForeColor = RGB(255, 0, 0) And Trim(Mt_desemp.TextMatrix(I, J)) <> "" Then
                notas_desemp.recuperado(J - 2) = True
            Else
                notas_desemp.recuperado(J - 2) = False
            End If
            notas_desemp.logro(J - 2) = codlogro(J - 2)
        Next J
        notas_desemp.num_carnet = Mt_desemp.TextMatrix(I, 1)
        Put #NAR, I, notas_desemp
    Next I
    Close #NAR
End If
VALI4 = True
Screen.MousePointer = 0
End Sub

Private Sub Command6_Click()
If Val(Mt_desemp.Rows - 1) = 0 Then
   MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 48, "IMPRIMIR"
   Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
   Screen.MousePointer = 11
   NAR = FreeFile
   Open Ruta & "inicial.edu" For Input As #NAR
   Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
   Close #NAR
   Printer.ScaleMode = 7
   Printer.Font.Size = 10
   Printer.CurrentY = 1
   Printer.CurrentX = 6.5
   'Printer.Print "PORCENTAJE DE LOGROS " & Frame1.Caption
   Printer.Print ""
   Printer.CurrentX = 0.5
   Printer.Print ini.nombre;
   Printer.CurrentX = 16.5
   Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
   Printer.CurrentX = 0.5
   Printer.Print Frame1.Caption
   Printer.Print ""
   Printer.CurrentX = 0.5
   Printer.Print "CD";
   Printer.CurrentX = 1.3
   Printer.Print "APELLIDOS Y NOMBRES";
   Printer.CurrentX = 10.5
   Printer.Print "PORCENTAJES DE LOGROS"
   Val_X = 10.5
   For I = 3 To Mt_desemp.Cols - 1
        Printer.CurrentX = Val_X
        Printer.Print "L-" & I - 2;
        Val_X = Val_X + 1
   Next I
   Printer.Print ""
    
   For I = 1 To (Mt_desemp.Rows - 1)
      Printer.CurrentX = 0.5
      Printer.Print I;
      Printer.CurrentX = 1.3
      Printer.Print RTrim(Mt_desemp.TextMatrix(I, 2));
      Val_X = 10.5
      For J = 3 To Mt_desemp.Cols - 1
        Printer.CurrentX = Val_X
        Printer.Print RTrim(Mt_desemp.TextMatrix(I, J));
        Val_X = Val_X + 1
      Next J
      Printer.Print ""
   Next I
   Printer.EndDoc
   Printer.Font.Size = 8
   Screen.MousePointer = 0
End If
End Sub

Private Sub Command7_Click()
RESP = MsgBox("Desea pegar los datos en toda la lista?", vbYesNo + vbQuestion + vbDefaultButton2, "Pegar todo")
If RESP = vbYes Then
    For vcpy = 1 To Val(Mt_desemp.Rows - 1)
        For kcpy = 3 To Val(Mt_desemp.Cols - 1)
            Mt_desemp.TextMatrix(vcpy, kcpy) = simucopy(kcpy)
        Next kcpy
    Next vcpy
End If
VALI4 = False
End Sub

Private Sub Command8_Click()
RESP = MsgBox("Desea pegar los datos en toda la columna?", vbYesNo + vbQuestion + vbDefaultButton2, "Pegar todo")
If RESP = vbYes Then
    If Mt_desemp.Col > 2 Then
        For vcpy = 1 To Val(Mt_desemp.Rows - 1)
            Mt_desemp.TextMatrix(vcpy, Mt_desemp.Col) = simucopy(Val(Mt_desemp.Col))
        Next vcpy
    End If
End If
VALI4 = False
End Sub

Private Sub Command9_Click()
Dim StudentNew As Boolean
I = 0
PASSW.Show 1
If I = 1 Then
    If Label4.Caption = "" Then
        MsgBox "Seleccione primero el grupo y el área y presione Ok", 48
        Exit Sub
    End If
    Screen.MousePointer = 11
    t = 1
    ret = 0
    NAR = FreeFile
    Open Ruta & Label4 & ".gru" For Random As #NAR Len = Len(alugru)
    While Not EOF(NAR)
        ret = ret + 1
        Get #NAR, ret, alugru
    Wend
    Close #NAR
    Open Ruta & Label4 & ".gru" For Random As #NAR Len = Len(alugru)
    For t = 1 To ret - 1
        StudentNew = False
        Get #NAR, t, alugru
        For s = 1 To Mt_desemp.Rows - 1
            If Val(alugru.num_carnet) = Val(Mt_desemp.TextMatrix(s, 1)) Then
                StudentNew = True
                Exit For
            End If
        Next s
        If StudentNew = False Then
            NAR = FreeFile
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, Val(alugru.num_carnet), alumno
            Close #NAR
            NAR = NAR - 1
            Mt_desemp.Rows = Mt_desemp.Rows + 1
            Mt_desemp.TextMatrix((Mt_desemp.Rows - 1), 2) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
            Mt_desemp.TextMatrix((Mt_desemp.Rows - 1), 1) = alumno.n_carnet
        End If
    Next t
    Close #NAR
    Mt_desemp.Col = 2
    Mt_desemp.Sort = 5
    For TT = 1 To Val(Mt_desemp.Rows - 1)
        Mt_desemp.TextMatrix(TT, 0) = TT
    Next TT
    VALI4 = False
    Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Load()
Mt_desemp.ColWidth(0) = 400
Mt_desemp.TextMatrix(0, 0) = "CD"
Mt_desemp.ColWidth(1) = 700
Mt_desemp.TextMatrix(0, 1) = "Carnet"
Mt_desemp.ColWidth(2) = 3800
Mt_desemp.TextMatrix(0, 2) = "APELLIDOS Y NOMBRES"
'Mt_desemp.CellForeColor = RGB(255, 255, 255)
'Mt_desemp.CellBackColor = RGB(0, 0, 150)
'Mt_desemp.Text = "OB1"
If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") Then
    Command4.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo2.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    cona = 0
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, mate
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    For I = 1 To cona - 1
        Get #NAR, I, mate
        If RTrim(mate.nom) <> "" Then
            Combo3.AddItem RTrim(mate.nom)
        End If
    Next I
    Close #NAR
    Combo2.Text = Combo2.List(0)
    Combo3.Text = Combo3.List(0)
Else
    Command4.Enabled = False
End If
Combo1 = Combo1.List(0)
Combo2 = Combo2.List(0)
Combo3 = Combo3.List(0)
Label6.Caption = ""
Label4.Caption = ""
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Check1.Enabled = False
VALI4 = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VALI4 = False Then
   Call Command5_Click
   Unload Me
Else
  Unload Me
End If
Unload Ver_Obser
End Sub

Private Sub Mt_desemp_Click()
If Mt_desemp.Col > 2 And Mt_desemp.Col < Mt_desemp.Cols Then
    If Check1.Value = 1 Then
        If Mt_desemp.CellForeColor = RGB(0, 0, 0) Then
            Mt_desemp.CellFontBold = True
            Mt_desemp.CellForeColor = RGB(255, 0, 0)
        Else
            Mt_desemp.CellFontBold = False
            Mt_desemp.CellForeColor = RGB(0, 0, 0)
        End If
    VALI4 = False
    End If
End If
End Sub

Private Sub Mt_desemp_KeyPress(KeyAscii As Integer)
If Mt_desemp.Col > 2 And Mt_desemp.Col < Mt_desemp.Cols Then
   If KeyAscii = 13 Then
      If Mt_desemp.Col = Mt_desemp.Cols - 1 And Mt_desemp.Row < Mt_desemp.Rows - 1 Then
         Mt_desemp.Row = Mt_desemp.Row + 1
         Mt_desemp.Col = 3
         Exit Sub
      End If
      If Mt_desemp.Col < Mt_desemp.Cols - 1 Then
        Mt_desemp.Col = Mt_desemp.Col + 1
      End If
      Exit Sub
   End If
   C$ = Chr(KeyAscii)
   If KeyAscii = 8 Then
      If Trim(Mt_desemp.Text) <> "" Then
         Mt_desemp.Text = Left(Mt_desemp.Text, Len(Mt_desemp.Text) - 1)
         If Mt_desemp.Text = "" Then
            Mt_desemp.CellForeColor = RGB(0, 0, 0)
         End If
         VALI4 = False
         Exit Sub
      Else
         If Mt_desemp.Col > 3 Then
            Mt_desemp.Col = Mt_desemp.Col - 1
         End If
      End If
   End If
   
   If C$ < "0" Or C$ > "9" Then
      KeyAscii = 0
      Beep
      Exit Sub
   End If
   rete = Chr(KeyAscii)
   Mt_desemp.CellForeColor = RGB(0, 0, 0)
   Mt_desemp.CellFontBold = False
   Mt_desemp.Text = Mt_desemp.Text + rete
   VALI4 = False
   If Val(Mt_desemp.Text) < 1 Or Val(Mt_desemp.Text) > 100 Then
        MsgBox "VALOR DE PORCENTAJE INVÁLIDO (Escriba un número de 1 a 100)", 48, "ADVERTENCIA"
        Mt_desemp.Text = ""
        Mt_desemp.CellForeColor = RGB(0, 0, 0)
        VALI4 = False
        Exit Sub
    End If
End If
End Sub
