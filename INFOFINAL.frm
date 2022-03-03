VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form INFREFINAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comentarios en el informe final"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "INFOFINAL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Height          =   320
      Left            =   2040
      Picture         =   "INFOFINAL.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Pegar observaciones creadas desde otra aplicación"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Height          =   320
      Left            =   6600
      Picture         =   "INFOFINAL.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir información mostrada en pantalla"
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Height          =   320
      Left            =   1440
      Picture         =   "INFOFINAL.frx":0A76
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Copiar observaciones existentes"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   320
      Left            =   240
      Picture         =   "INFOFINAL.frx":0B78
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Agregar observación"
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   320
      Left            =   6000
      Picture         =   "INFOFINAL.frx":0C7A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Guardar lista del grupo"
      Top             =   5760
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   320
      Left            =   2640
      TabIndex        =   2
      Top             =   5760
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comentarios finales"
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   6975
      Begin VB.CommandButton Command7 
         Height          =   320
         Left            =   720
         Picture         =   "INFOFINAL.frx":11AC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminar observación"
         Top             =   2280
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   2010
         ItemData        =   "INFOFINAL.frx":12AE
         Left            =   120
         List            =   "INFOFINAL.frx":12B0
         TabIndex        =   11
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin MSFlexGridLib.MSFlexGrid MATIRF 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   2
         BackColorBkg    =   12632256
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GRUPO:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   630
   End
End
Attribute VB_Name = "INFREFINAL"
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
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
'Dim obsfin As String * 120
'Dim leyfin As leyenfin
'Dim alumno As maestroalum
'Dim alugru As grupoalu
If VALI = False Then
    Call Command2_Click
End If
Frame1.Caption = ""
List1.ToolTipText = ""
MATIRF.ToolTipText = ""
MATIRF.Rows = 1
List1.Clear
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "Grupo incorrecto"
    Exit Sub
End If
Frame1.Caption = Combo1.Text
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
NAR = FreeFile
If Dir(Ruta & "LRF" & Frame1.Caption & ".lrf") <> "" Then
    Command6.Enabled = False
    cona = 0
    Open Ruta & "LRF" & Frame1.Caption & ".lrf" For Random As #NAR Len = Len(leyfin)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, leyfin
    Wend
    Close #NAR
    Open Ruta & "LRF" & Frame1.Caption & ".lrf" For Random As #NAR Len = Len(leyfin)
    For I = 1 To (cona - 1)
        Get #NAR, I, leyfin
        MATIRF.Rows = MATIRF.Rows + 1
        MATIRF.TextMatrix((MATIRF.Rows - 1), 0) = I
        NAR = FreeFile
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(leyfin.num_carnet)), alumno
        Close #NAR
        NAR = NAR - 1
        MATIRF.TextMatrix((MATIRF.Rows - 1), 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        If leyfin.fnob(1) <> 0 Then MATIRF.TextMatrix((MATIRF.Rows - 1), 2) = leyfin.fnob(1)
        If leyfin.fnob(2) <> 0 Then MATIRF.TextMatrix((MATIRF.Rows - 1), 3) = leyfin.fnob(2)
        If leyfin.fnob(3) <> 0 Then MATIRF.TextMatrix((MATIRF.Rows - 1), 4) = leyfin.fnob(3)
        If leyfin.fnob(4) <> 0 Then MATIRF.TextMatrix((MATIRF.Rows - 1), 5) = leyfin.fnob(4)
        If leyfin.fnob(5) <> 0 Then MATIRF.TextMatrix((MATIRF.Rows - 1), 6) = leyfin.fnob(5)
        MATIRF.TextMatrix((MATIRF.Rows - 1), 7) = leyfin.num_carnet
    Next I
    Close #NAR
Else
    cona = 0
    Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, alugru
    Wend
    Close #NAR
    Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
    For I = 1 To (cona - 1)
        Get #NAR, I, alugru
        MATIRF.Rows = MATIRF.Rows + 1
        MATIRF.TextMatrix((MATIRF.Rows - 1), 0) = I
        NAR = FreeFile
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        NAR = NAR - 1
        MATIRF.TextMatrix((MATIRF.Rows - 1), 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        MATIRF.TextMatrix((MATIRF.Rows - 1), 7) = alugru.num_carnet
    Next I
    Close #NAR
End If
If Dir(Ruta & "ORF" & Frame1.Caption & ".orf") <> "" Then
    cona = 0
    Open Ruta & "ORF" & Frame1.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, obsfin
    Wend
    Close #NAR
    Open Ruta & "ORF" & Frame1.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
    For I = 1 To (cona - 1)
        Get #NAR, I, obsfin
        List1.AddItem RTrim(obsfin)
    Next I
    Close #NAR
End If
End Sub

Private Sub Command2_Click()
'Dim leyfin As leyenfin
If Frame1.Caption = "" Then
    MsgBox "No existe información para Guardar", 64, "Guardar"
    Exit Sub
End If
RESP = MsgBox("Desea Guardar la información?", vbYesNo + vbQuestion + vbDefaultButton1, "Guardar")
If RESP = vbYes Then
    Y = 0
    Screen.MousePointer = 11
    If Dir(Ruta & "LRF" & Frame1.Caption & ".lrf") <> "" Then
        Kill Ruta & "LRF" & Frame1.Caption & ".lrf"
    End If
    NAR = FreeFile
    Open Ruta & "LRF" & Frame1.Caption & ".lrf" For Random As #NAR Len = Len(leyfin)
    For I = 1 To (MATIRF.Rows - 1)
        For J = 2 To 6
            If MATIRF.TextMatrix(I, J) <> "" Then
                leyfin.fnob(J - 1) = MATIRF.TextMatrix(I, J)
                Y = 1
            Else
                leyfin.fnob(J - 1) = 0
            End If
        Next J
        leyfin.num_carnet = MATIRF.TextMatrix(I, 7)
        Put #NAR, I, leyfin
    Next I
    Close #NAR
    If Y = 0 Then
        If Dir(Ruta & "LRF" & Frame1.Caption & ".lrf") <> "" Then
            Kill Ruta & "LRF" & Frame1.Caption & ".lrf"
        End If
    End If
End If
VALI = True
Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
'Dim obsfin As String * 120
If Frame1.Caption = "" Then
    MsgBox "Seleccione primero un grupo"
    Exit Sub
End If
If List1.ListCount = 10 Then
    MsgBox "No se pueden agregar más observaciones", 48, "Observaciones"
Else
    TTT = InputBox("Observación No." & (List1.ListCount + 1), "Agregar observación")
    If RTrim(TTT) <> "" Then
        NAR = FreeFile
        Open Ruta & "ORF" & Frame1.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
        obsfin = TTT
        Put #NAR, (List1.ListCount + 1), obsfin
        Close #NAR
        List1.AddItem TTT
    End If
End If
End Sub

Private Sub Command4_Click()
If (List1.ListCount) = 0 Then
    MsgBox "No existen observaciones para copiar", 64, "Copiar"
    Exit Sub
End If
Clipboard.Clear
cop = ""
For X = 0 To (List1.ListCount - 1)
    cop = cop + List1.List(X) & vbCrLf
Next X
Clipboard.SetText cop
End Sub

Private Sub Command5_Click()
'Dim ini As inicio
If (List1.ListCount > 0) And (Frame1.Caption <> "") Then
    RESP = MsgBox("Desea imprimir la información?", vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir")
        If RESP = vbYes Then
            Screen.MousePointer = 11
            NAR = FreeFile
            Open Ruta & "inicial.edu" For Input As #NAR
            Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
            Close #NAR
            Printer.ScaleMode = 7
            Printer.CurrentY = 1
            Printer.CurrentX = 7
            Printer.Print "INFORMACION INFORME FINAL - " & Frame1.Caption
            Printer.Print ""
            Printer.CurrentX = 2
            Printer.Print ini.nombre
            Printer.Print ""
            Printer.Print ""
            Printer.CurrentX = 2
            Printer.Print "CD";
            Printer.CurrentX = 2.6
            Printer.Print "CARNET";
            Printer.CurrentX = 4
            Printer.Print "APELLIDOS Y NOMBRES";
            Printer.CurrentX = 12
            Printer.Print "OB1";
            Printer.CurrentX = 13
            Printer.Print "OB2";
            Printer.CurrentX = 14
            Printer.Print "OB3";
            Printer.CurrentX = 15
            Printer.Print "OB4";
            Printer.CurrentX = 16
            Printer.Print "OB5"
            Printer.Print ""
            For I = 1 To (MATIRF.Rows - 1)
                Printer.CurrentX = 2
                Printer.Print MATIRF.TextMatrix(I, 0);
                Printer.CurrentX = 2.6
                Printer.Print MATIRF.TextMatrix(I, 7);
                Printer.CurrentX = 4
                Printer.Print MATIRF.TextMatrix(I, 1);
                Printer.CurrentX = 12.2
                Printer.Print MATIRF.TextMatrix(I, 2);
                Printer.CurrentX = 13.2
                Printer.Print MATIRF.TextMatrix(I, 3);
                Printer.CurrentX = 14.2
                Printer.Print MATIRF.TextMatrix(I, 4);
                Printer.CurrentX = 15.2
                Printer.Print MATIRF.TextMatrix(I, 5);
                Printer.CurrentX = 16.2
                Printer.Print MATIRF.TextMatrix(I, 6)
            Next I
            Printer.CurrentY = 21.5
            Printer.Line (2, Printer.CurrentY)-(20.2, Printer.CurrentY)
            Printer.Print ""
            Printer.CurrentX = 2
            Printer.Print "OBSERVACIONES FINALES:"
            Printer.Print ""
            For I = 1 To List1.ListCount
                Printer.CurrentX = 2
                Printer.Print I & "- " & RTrim(List1.List(I - 1))
            Next I
            Printer.EndDoc
            Screen.MousePointer = 0
        End If
Else
        MsgBox "No existe información para imprimir", 64, "Imprimir"
End If
End Sub

Private Sub Command6_Click()
'Dim obsfin As String * 120
If Frame1.Caption = "" Then Exit Sub
peg = Clipboard.GetText(1)
If peg = "" Then Exit Sub
If List1.ListCount = 0 Then
    GoTo pegar
End If
RESP = MsgBox("Desea reemplazar las observaciones existentes y pegar unas nuevas?", vbYesNo + vbQuestion + vbDefaultButton2, "Pegar observaciones")
If RESP = vbYes Then
pegar:
    List1.Clear
    d = 0
    Last = 1
    If Dir(Ruta & "ORF" & Frame1.Caption & ".orf") <> "" Then
        Kill Ruta & "ORF" & Frame1.Caption & ".orf"
    End If
    For X = 1 To Len(peg)
        mpeg = Mid(peg, X, 1)
        smpeg = Mid(peg, X + 1, 1)
        If mpeg = vbCr Or mpeg = vbCrLf Or mpeg = vbLf Then
            If (X = 1) And (Last = 1) Then Exit Sub
            PEGG = Mid(peg, Last, X - Last - 1)
            If RTrim(PEGG) <> "" Then
                NAR = FreeFile
                Open Ruta & "ORF" & Frame1.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
                obsfin = LTrim(Mid(PEGG, 1, Len(PEGG)))
                Put #NAR, (d + 1), obsfin
                Close #NAR
                List1.AddItem RTrim(LTrim(Mid(PEGG, 1, Len(PEGG))))
                d = d + 1
                If d > 9 Then Exit Sub
            End If
            Last = X + 1
        End If
        If smpeg = vbCr Or smpeg = vbCrLf Or smpeg = vbLf Then
            X = X + 1
        End If
    Next X
    If Last <= Len(peg) Then
        PEGG = Mid(peg, Last, Len(peg))
        If RTrim(PEGG) <> "" Then
            NAR = FreeFile
            Open Ruta & "ORF" & Frame1.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
            obsfin = LTrim(Mid(PEGG, 1, Len(PEGG)))
            Put #NAR, (d + 1), obsfin
            Close #NAR
            List1.AddItem RTrim(LTrim(Mid(PEGG, 1, Len(PEGG))))
        End If
    End If
Frame2.Caption = "Observaciones finales"
For J = 1 To (MATIRF.Rows - 1)
    For I = 2 To 6
        MATIRF.TextMatrix(J, I) = ""
    Next I
Next J
If Dir(Ruta & "LRF" & Frame1.Caption & ".lrf") <> "" Then
    Kill Ruta & "LRF" & Frame1.Caption & ".lrf"
End If
VALI = True
End If
End Sub

Private Sub Command7_Click()
'Dim leyfin As leyenfin
'Dim obsfin As String * 120
If Frame1.Caption = "" Then Exit Sub
If List1.ListCount > 0 Then
    If List1.ListIndex > -1 Then
        RESP = MsgBox("Desea eliminar esta observación?", vbYesNo + vbQuestion + vbDefaultButton2, "Eliminar observación No." & (List1.ListIndex + 1))
        If RESP = vbYes Then
            Y = 0
            X = List1.ListIndex + 1
            List1.RemoveItem (List1.ListIndex)
            Kill Ruta & "ORF" & Frame1.Caption & ".orf"
            NAR = FreeFile
            For I = 1 To List1.ListCount
                Open Ruta & "ORF" & Frame1.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
                obsfin = List1.List(I - 1)
                Put #NAR, I, obsfin
                Close #NAR
            Next I
            Frame2.Caption = "Observaciones finales"
            For J = 1 To (MATIRF.Rows - 1)
                For I = 2 To 6
                    If Val(MATIRF.TextMatrix(J, I)) = X Then
                        MATIRF.TextMatrix(J, I) = ""
                    End If
                    If Val(MATIRF.TextMatrix(J, I)) > X Then
                        MATIRF.TextMatrix(J, I) = Val(MATIRF.TextMatrix(J, I)) - 1
                    End If
                Next I
            Next J
            If Dir(Ruta & "LRF" & Frame1.Caption & ".lrf") <> "" Then
                Kill Ruta & "LRF" & Frame1.Caption & ".lrf"
            End If
            If List1.ListCount > 0 Then
                NAR = FreeFile
                Open Ruta & "LRF" & Frame1.Caption & ".lrf" For Random As #NAR Len = Len(leyfin)
                For I = 1 To (MATIRF.Rows - 1)
                    For J = 2 To 6
                        If MATIRF.TextMatrix(I, J) <> "" Then
                            leyfin.fnob(J - 1) = MATIRF.TextMatrix(I, J)
                            Y = 1
                        Else
                            leyfin.fnob(J - 1) = 0
                        End If
                    Next J
                    leyfin.num_carnet = MATIRF.TextMatrix(I, 7)
                    Put #NAR, I, leyfin
                Next I
                Close #NAR
                If Y = 0 Then
                    If Dir(Ruta & "LRF" & Frame1.Caption & ".lrf") <> "" Then
                        Kill Ruta & "LRF" & Frame1.Caption & ".lrf"
                    End If
                End If
            Else
                For J = 1 To (MATIRF.Rows - 1)
                    For I = 2 To 6
                        MATIRF.TextMatrix(J, I) = ""
                    Next I
                Next J
            End If
            VALI = True
        End If
    Else
        MsgBox "Seleccione la observación que desea eliminar", 64, "Eliminar"
    End If
Else
    MsgBox "No existen observaciones para eliminar", 64, "Eliminar"
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Comentarios en el informe final: Esta información saldrá en cada informe final de acuerdo al grupo."
End Sub

Private Sub Form_Load()
'Dim icur As inforcur
MATIRF.Row = 0
MATIRF.Col = 0
MATIRF.ColWidth(0) = 350
MATIRF.CellForeColor = RGB(255, 255, 255)
MATIRF.CellBackColor = RGB(0, 0, 150)
MATIRF.Text = "CD"
MATIRF.Col = 1
MATIRF.ColWidth(1) = 4000
MATIRF.CellForeColor = RGB(255, 255, 255)
MATIRF.CellBackColor = RGB(0, 0, 150)
MATIRF.Text = "APELLIDOS Y NOMBRES"
MATIRF.Col = 2
MATIRF.ColWidth(2) = 400
MATIRF.CellForeColor = RGB(255, 255, 255)
MATIRF.CellBackColor = RGB(0, 0, 150)
MATIRF.Text = "OB1"
MATIRF.Col = 3
MATIRF.ColWidth(3) = 400
MATIRF.CellForeColor = RGB(255, 255, 255)
MATIRF.CellBackColor = RGB(0, 0, 150)
MATIRF.Text = "OB2"
MATIRF.Col = 4
MATIRF.ColWidth(4) = 400
MATIRF.CellForeColor = RGB(255, 255, 255)
MATIRF.CellBackColor = RGB(0, 0, 150)
MATIRF.Text = "OB3"
MATIRF.Col = 5
MATIRF.ColWidth(5) = 400
MATIRF.CellForeColor = RGB(255, 255, 255)
MATIRF.CellBackColor = RGB(0, 0, 150)
MATIRF.Text = "OB4"
MATIRF.Col = 6
MATIRF.ColWidth(6) = 400
MATIRF.CellForeColor = RGB(255, 255, 255)
MATIRF.CellBackColor = RGB(0, 0, 150)
MATIRF.Text = "OB5"
MATIRF.Col = 7
MATIRF.ColWidth(7) = 1000
MATIRF.CellForeColor = RGB(255, 255, 255)
MATIRF.CellBackColor = RGB(0, 0, 150)
MATIRF.Text = "CARNET"
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
VALI = True
If Dir(Ruta & "infcur.edu") <> "" Then
    Command1.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo1.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Combo1.Text = Combo1.List(0)
Else
    Command1.Enabled = False
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VALI = False Then
   Call Command2_Click
   Unload Me
Else
   Unload Me
End If
End Sub

Private Sub List1_Click()
List1.ToolTipText = "[" & (List1.ListIndex + 1) & "] " & Mid(List1.List(List1.ListIndex), 1, 200)
Frame2.Caption = "Observaciones finales - No." & (List1.ListIndex + 1)
End Sub

Private Sub List1_DblClick()
'Dim obsfin As String * 120
TTT = List1.List(List1.ListIndex)
TTT = InputBox("Observación No." & (List1.ListIndex + 1), "Corregir observación", TTT)
If RTrim(TTT) <> "" Then
    NAR = FreeFile
    Open Ruta & "ORF" & Frame1.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
    obsfin = TTT
    Put #NAR, (List1.ListIndex + 1), obsfin
    Close #NAR
    List1.List(List1.ListIndex) = TTT
End If
End Sub

Private Sub MATIRF_Click()
MATIRF.ToolTipText = ""
If MATIRF.Col > 1 And MATIRF.Col < 7 And MATIRF.Row > 0 Then
    If (MATIRF.Text <> "") And (Val(MATIRF.Text) <= List1.ListCount) Then
        MATIRF.ToolTipText = List1.List(Val(MATIRF.Text) - 1)
    End If
End If
End Sub

Private Sub MATIRF_KeyPress(KeyAscii As Integer)
If MATIRF.Col < 7 And MATIRF.Col > 1 And MATIRF.Row > 0 And MATIRF.Row <= (MATIRF.Rows - 1) Then
   If KeyAscii = 13 Then
      If (MATIRF.Col = 6) And (MATIRF.Row <> (MATIRF.Rows - 1)) Then
         MATIRF.Row = MATIRF.Row + 1
         MATIRF.Col = 2
         Exit Sub
      End If
      MATIRF.Col = MATIRF.Col + 1
      Exit Sub
   End If
   C$ = Chr(KeyAscii)
   If KeyAscii = 8 Then
      If MATIRF.Text <> "" Then
         MATIRF.Text = Left(MATIRF.Text, Len(MATIRF.Text) - 1)
         VALI = False
         Exit Sub
      Else
         If MATIRF.Col > 2 Then
            MATIRF.Col = MATIRF.Col - 1
         End If
      End If
   End If
   If C$ < "0" Or C$ > "9" Then
      KeyAscii = 0
      Beep
      Exit Sub
   End If
   rete = Chr(KeyAscii)
   MATIRF.CellFontBold = True
   MATIRF.CellForeColor = RGB(0, 0, 255)
   MATIRF.Text = MATIRF.Text + rete
   VALI = False
   If (Val(MATIRF.Text) < 1) Or (Val(MATIRF.Text) > List1.ListCount) Then
        MsgBox "Observación no existe", 64, "Información informe final"
        MATIRF.Text = ""
   End If
End If
End Sub
