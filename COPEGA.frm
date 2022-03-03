VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form COPEGA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logros y observaciones por grado"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "COPEGA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Eliminar una observación de la lista de observaciones"
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      ToolTipText     =   "Imprimir las observaciones existentes"
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Agregar una nueva observación"
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      ToolTipText     =   "Guardar las observaciones existentes"
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Pegar"
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      ToolTipText     =   "Pegar observaciones creadas desde otra aplicación"
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Copiar"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "Copiar las observaciones existentes"
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00C00000&
      Height          =   6015
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9255
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "COPEGA.frx":0442
         Left            =   7680
         List            =   "COPEGA.frx":0455
         TabIndex        =   0
         Text            =   "PRIMERO"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "OPCIONES"
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   4560
         Width           =   9015
         Begin VB.CommandButton Command5 
            Caption         =   "&OK"
            Height          =   320
            Left            =   8160
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   4920
            TabIndex        =   3
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "COPEGA.frx":0483
            Left            =   3000
            List            =   "COPEGA.frx":04B1
            TabIndex        =   2
            Text            =   "PREJARDIN"
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "COPEGA.frx":0531
            Left            =   960
            List            =   "COPEGA.frx":0541
            TabIndex        =   1
            Text            =   "UNICA"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "AREA:"
            Height          =   195
            Left            =   4440
            TabIndex        =   25
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "GRADO:"
            Height          =   195
            Left            =   2280
            TabIndex        =   24
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "JORNADA:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   810
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   5280
         Width           =   9015
         Begin VB.ComboBox Combo5 
            Height          =   315
            ItemData        =   "COPEGA.frx":0562
            Left            =   600
            List            =   "COPEGA.frx":0572
            TabIndex        =   6
            Text            =   "L"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command6 
            Height          =   300
            Left            =   8520
            Picture         =   "COPEGA.frx":0582
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Subir la observación a la lista de observaciones"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   240
            Width           =   7095
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MATU20 
         Height          =   4095
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PERIODO:"
         Height          =   195
         Left            =   6840
         TabIndex        =   22
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8040
      TabIndex        =   19
      Top             =   6360
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7680
      TabIndex        =   18
      Top             =   6360
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "COPEGA"
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
    Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo3.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo4.SetFocus
End If
End Sub

Private Sub Combo4_Change()
If Combo4.Text <> Combo4.List(0) Then
    Combo4.Text = Combo4.List(0)
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If Command5.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command5_Click
End If
End Sub

Private Sub Command1_Click()
If Val(Label1.Caption) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA COPIAR", 16, "COPIAR"
    Exit Sub
End If
Screen.MousePointer = 11
Clipboard.Clear
cop = ""
For X = 1 To Val(Label1.Caption)
    INDI = RTrim(MATU20.TextMatrix(X, 1))
    OB = RTrim(MATU20.TextMatrix(X, 2))
    If Trim(INDI) <> "" Then
        cop = cop + INDI & " - " & OB & vbCrLf
    Else
        cop = cop + INDI & OB & vbCrLf
    End If
Next X
Close #NAR
Clipboard.SetText cop
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
peg = Clipboard.GetText(1)
If peg = "" Then Exit Sub
If Val(Label1.Caption) = 0 Then
    GoTo pepega
End If
RESP = MsgBox("Desea reemplazar las observaciones existentes y pegar unas nuevas?", vbYesNo + vbQuestion + vbDefaultButton2, "Pegar observaciones")
If RESP = vbYes Then
pepega:
MATU20.Rows = 1
Text1.Text = ""
Text3.Text = ""
Last = 1
d = 1
For X = 1 To Len(peg)
    mpeg = Mid(peg, X, 1)
    smpeg = Mid(peg, X + 1, 1)
    If mpeg = vbCr Or mpeg = vbCrLf Or mpeg = vbLf Then
      MATU20.Rows = d + 1
      MATU20.TextMatrix(d, 0) = d
      If (X = 1) And (Last = 1) Then Exit Sub
      PEGG = Mid(peg, Last, X - Last - 1)
'      If (Mid(PEGG, 3, 1) = "-") And (Mid(PEGG, 2, 1) = " ") And (Mid(PEGG, 4, 1) = " ") Then
      If Mid(PEGG, 2, 3) = " - " Then
        MATU20.TextMatrix(d, 1) = Mid(PEGG, 1, 1)
        MATU20.TextMatrix(d, 2) = LTrim(Mid(PEGG, 5, Len(PEGG)))
      Else
        If Mid(PEGG, 3, 3) = " - " Then
            MATU20.TextMatrix(d, 1) = Mid(PEGG, 1, 2)
            MATU20.TextMatrix(d, 2) = LTrim(Mid(PEGG, 6, Len(PEGG)))
        Else
            If Mid(PEGG, 4, 3) = " - " Then
                MATU20.TextMatrix(d, 1) = Mid(PEGG, 1, 3)
                MATU20.TextMatrix(d, 2) = LTrim(Mid(PEGG, 7, Len(PEGG)))
            Else
                MATU20.TextMatrix(d, 2) = LTrim(Mid(PEGG, 1, Len(PEGG)))
            End If
        End If
      End If
      Last = X + 1
      d = d + 1
    End If
    If smpeg = vbCr Or smpeg = vbCrLf Or smpeg = vbLf Then
      X = X + 1
    End If
Next X
    If Last <= Len(peg) Then
        MATU20.Rows = d + 1
        MATU20.TextMatrix(d, 0) = d
        PEGG = Mid(peg, Last, Len(peg))
        If Mid(PEGG, 2, 3) = " - " Then
          MATU20.TextMatrix(d, 1) = Mid(PEGG, 1, 1)
          MATU20.TextMatrix(d, 2) = LTrim(Mid(PEGG, 5, Len(PEGG)))
        Else
          If Mid(PEGG, 3, 3) = " - " Then
              MATU20.TextMatrix(d, 1) = Mid(PEGG, 1, 2)
              MATU20.TextMatrix(d, 2) = LTrim(Mid(PEGG, 6, Len(PEGG)))
          Else
              If Mid(PEGG, 4, 3) = " - " Then
                  MATU20.TextMatrix(d, 1) = Mid(PEGG, 1, 3)
                  MATU20.TextMatrix(d, 2) = LTrim(Mid(PEGG, 7, Len(PEGG)))
              Else
                  MATU20.TextMatrix(d, 2) = LTrim(Mid(PEGG, 1, Len(PEGG)))
              End If
          End If
        End If
    Else
    d = d - 1
    End If
Label1.Caption = d
VALI80 = False
End If
End Sub

Private Sub Command3_Click()
If VALI80 = False Then
   Call Command4_Click
   Unload Me
Else
   Unload Me
End If
End Sub

Private Sub Command4_Click()
If Val(Label1.Caption) = 0 Then
    MsgBox "NO HAY INFORMACION PARA GUARDAR", 64
    Exit Sub
End If
    MS1 = "DESEA GUARDAR ESTAS OBSERVACIONES?"
    If FileLen(Label2.Caption) <> 0 Then
       MS1 = "DESEA GUARDAR LOS CAMBIOS EFECTUADOS?"
    End If
    RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
    If RESP = vbYes Then
       Screen.MousePointer = 11
       Kill Label2.Caption
       NAR = FreeFile
       Open Label2.Caption For Random As #NAR Len = Len(logru)
       For X = 1 To Val(Label1.Caption)
           logru.indicador = Format(MATU20.TextMatrix(X, 1), ">")
           logru.observ = RTrim(MATU20.TextMatrix(X, 2))
           Put #NAR, X, logru
       Next X
       Close #NAR
       Screen.MousePointer = 0
    End If
    VALI80 = True
End Sub

Private Sub Command5_Click()
Dim ConfLgr As Byte
If VALI80 = False Then
   Call Command4_Click
End If
MATU20.ToolTipText = ""
Y = 0
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    Y = Y + 1
    Get #NAR, Y, mate
    If RTrim(mate.nom) = Combo4.Text Then
        NA = mate.num
    End If
Wend
Close #NAR
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If (RTrim(argra.grado) = RTrim(Combo3.Text) And (argra.num_area = NA)) Then
        Close #NAR
        GoTo intel
    End If
Wend
Close #NAR
MsgBox "ESTA MATERIA NO ESTA CREADA PARA ESTE GRADO", 16, "OBSERVACIONES"
Frame1.Caption = ""
MATU20.Rows = 1
Text1.Text = ""
Text3.Text = ""
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Exit Sub
intel:
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
If Combo2.Text = "UNICA" Then
fl = "1"
End If
If Combo2.Text = "MAÑANA" Then
fl = "2"
End If
If Combo2.Text = "TARDE" Then
fl = "3"
End If
If Combo2.Text = "NOCHE" Then
fl = "4"
End If
ser = Left(Combo3.Text, 3)
MATU20.Rows = 1
CROA = 0
Open Ruta & fl & ser & NA & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    CROA = CROA + 1
    Get #NAR, CROA, logru
Wend
Close #NAR
Label1.Caption = CROA - 1
Open Ruta & fl & ser & NA & lw & ".lgr" For Random As #NAR Len = Len(logru)
For J = 1 To Val(Label1.Caption)
    Get #NAR, J, logru
    MATU20.Rows = J + 1
    MATU20.TextMatrix(J, 0) = J
    MATU20.TextMatrix(J, 1) = logru.indicador
    MATU20.TextMatrix(J, 2) = logru.observ
Next J
Close #NAR
Text1.Text = ""
Text3.Text = ""
Label2.Caption = Ruta & fl & ser & NA & lw & ".lgr"
Frame1.Caption = "JORNADA: " & Combo2.Text & " - " & "GRADO: " & Combo3.Text & " - " & "AREA: " & Combo4.Text & " (" & NA & ")" & " - " & "PERIODO: " & Combo1.Text
Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
'Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True

'NO SE PUEDE ELIMINAR NI PEGAR INFORMACIÓN SI EXISTEN PORCENTAJES DE LOGROS GRABADOS
If Dir(Ruta & "conf_logro.edu") <> "" Then
    Open Ruta & "conf_logro.edu" For Input As #NAR
    Input #NAR, ConfLgr
    Close #NAR
    If ConfLgr = 1 Then
        If Dir(Ruta & fl & ser & NA & lw & ".ptj") <> "" Then
            Command2.Enabled = False
            Command9.Enabled = False
            Exit Sub
        End If
    End If
End If
' NO SE PUEDE ELIMINAR NI PEGAR INFORMACIÓN SI EXISTEN OBSERVACIONES O DESEMPEÑOS GRABADOS.
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If (RTrim(argra.grado) = RTrim(Combo3.Text) And (argra.num_area = NA)) Then
        If (Dir(Ruta & RTrim(argra.nom_grup) & NA & lw & ".obs") <> "") Or (Dir(Ruta & RTrim(argra.nom_grup) & NA & lw & ".dsp") <> "") Then
            Close #NAR
            Command2.Enabled = False
            Command9.Enabled = False
            Exit Sub
        End If
    End If
Wend
Close #NAR
End Sub

Private Sub Command6_Click()
Dim ConfLgr As Byte
If Text1.Text = "" Then
    MsgBox "NO HAY INFORMACION PARA SUBIR", 48
    Exit Sub
End If
If (Format(Trim(MATU20.TextMatrix(Val(Text1.Text), 1)), ">") = "L") Or (Format(Trim(Combo5.Text), ">") = "L") Then
    If Trim(MATU20.TextMatrix(Val(Text1.Text), 1)) <> Trim(Combo5.Text) Then
        'NO SE PUEDEN MODIFICAR LOGROS SI EXISTEN GRABADOS PORCENTAJES MANUALES DE LOGROS
        If Dir(Ruta & "conf_logro.edu") <> "" Then
            Open Ruta & "conf_logro.edu" For Input As #NAR
            Input #NAR, ConfLgr
            Close #NAR
            If ConfLgr = 1 Then
                If Dir(Ruta & fl & ser & NA & lw & ".ptj") <> "" Then
                    MsgBox "No se puede modificar el indicador (L), ni agregar nuevos logros, ya que existen porcentajes de logros para esta materia (deberá borrar los porcentajes de este grado para hacer cambios)", 16
                    Exit Sub
                End If
            End If
        End If
        ' NO SE PUEDEN MODIFICAR LOGROS SI EXISTEN DESEMPEÑOS GRABADOS
        NAR = FreeFile
        cona = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            cona = cona + 1
            Get #NAR, cona, argra
            If (RTrim(argra.grado) = RTrim(Combo3.Text) And (argra.num_area = NA)) Then
                If Dir(Ruta & RTrim(argra.nom_grup) & NA & lw & ".dsp") <> "" Then
                    Close #NAR
                    MATU20.TextMatrix(Val(Text1.Text), 2) = Text3.Text
                    MsgBox "No se puede modificar el indicador (L), ni agregar nuevos logros, ya que existen notas para esta materia (deberá borrar las notas de este grado para hacer cambios)", 16
                    Exit Sub
                End If
            End If
        Wend
        Close #NAR
    End If
End If
MATU20.TextMatrix(Val(Text1.Text), 1) = Trim(Format(Left(Combo5.Text, 3), ">"))
MATU20.TextMatrix(Val(Text1.Text), 2) = Text3.Text
VALI80 = False
End Sub

Private Sub Command7_Click()
'Text3.Enabled = True
'Combo5.Enabled = True
Command3.Enabled = True
'Command6.Enabled = True

Label1.Caption = Val(Label1.Caption) + 1
MATU20.Rows = Val(Label1.Caption) + 1
MATU20.TextMatrix(Val(Label1.Caption), 0) = Val(Label1.Caption)
Text1.Text = Val(Label1.Caption)
Text3.Text = ""
Combo5.SetFocus
End Sub

Private Sub Command8_Click()
ImprimirObserv.Show 1
End Sub

Private Sub Command9_Click()
If Val(Label1.Caption) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA ELIMINAR", 64
    Exit Sub
End If
If Val(Label1.Caption) = 1 Then
    MsgBox "No se puede Eliminar la última observación de la lista", 32, "Eliminar observación"
    Exit Sub
End If
TTT = InputBox("Escriba el número de la observación que desea eliminar (de 1 a " & Label1.Caption & ")", "Eliminar observación")
If TTT = "" Then
    MsgBox "No escribió el No. de la observación", 64, "Eliminar observación"
    Exit Sub
End If
If (Val(TTT) > Val(Label1.Caption)) Or (Val(TTT) < 1) Then
    MsgBox "Número de observación no existe", 64, "Eliminar observación"
    Exit Sub
End If
MATU20.RemoveItem Val(TTT)
Label1.Caption = Val(Label1.Caption) - 1
For TT = 1 To Val(MATU20.Rows - 1)
    MATU20.TextMatrix(TT, 0) = TT
Next TT
'Text1.Text = Val(Text1.Text) - 1
Text1.Text = ""
Text3.Text = ""
VALI80 = False
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta y grabación de observaciones por grado."
End Sub

Private Sub Form_Load()
'Dim mate As infomater
MATU20.Row = 0
MATU20.Col = 0
MATU20.ColWidth(0) = 400
MATU20.CellForeColor = RGB(255, 255, 255)
MATU20.CellBackColor = RGB(0, 0, 150)
MATU20.Text = "No."
MATU20.Col = 1
MATU20.ColWidth(1) = 400
MATU20.CellForeColor = RGB(255, 255, 255)
MATU20.CellBackColor = RGB(0, 0, 150)
MATU20.Text = "IND"
MATU20.Col = 2
MATU20.ColWidth(2) = 10250
MATU20.CellForeColor = RGB(255, 255, 255)
MATU20.CellBackColor = RGB(0, 0, 150)
MATU20.Text = "OBSERVACION"
If Dir(Ruta & "materia.edu") <> "" Then
    Command5.Enabled = True
    NAR = FreeFile
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    que = 0
    While Not EOF(NAR)
        que = que + 1
        Get #NAR, que, mate
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    For I = 1 To que - 1
        Get #NAR, I, mate
        If RTrim(mate.nom) <> "" Then
            Combo4.AddItem RTrim(mate.nom)
        End If
    Next I
    Close #NAR
    Combo4.Text = Combo4.List(0)
Else
    Command5.Enabled = False
End If
Text1.Text = ""
Text3.Text = ""
Text3.MaxLength = 800
Text3.Enabled = False
Combo5.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Label1.Caption = 0
VALI80 = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VALI80 = False Then
   Call Command4_Click
   Unload Me
Else
   Unload Me
End If
End Sub

Private Sub MATU20_Click()
If MATU20.Row > 0 Then
    MATU20.Col = 0
    Text1.Text = MATU20.Text
    MATU20.Col = 1
    Combo5.Text = RTrim(MATU20.Text)
    MATU20.Col = 2
    Text3.Text = RTrim(MATU20.Text)
    Text3.SetFocus
    MATU20.ToolTipText = Left(RTrim(MATU20.Text), 200)
End If
End Sub
Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ >= "0" And C$ <= "9" Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
    Text3.Enabled = True
    Combo5.Enabled = True
    Command6.Enabled = True
Else
    Text3.Enabled = False
    Combo5.Enabled = False
    Command6.Enabled = False
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Command6.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command6_Click
End If
End Sub
