VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form LOGROS_REPORT_AREA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentaje de logros por área"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9375
   Icon            =   "LOGROS_REPORT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   9135
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   6600
         TabIndex        =   10
         Top             =   120
         Width           =   2415
         Begin VB.CommandButton Command2 
            Caption         =   "&IMPRIMIR"
            Height          =   735
            Left            =   600
            Picture         =   "LOGROS_REPORT.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir reporte"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6375
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   4320
            TabIndex        =   3
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   960
            TabIndex        =   2
            Top             =   600
            Width           =   2535
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   960
            TabIndex        =   1
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "AREA   :"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO:"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   630
         End
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid MATI50 
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   2
         BackColorBkg    =   12632256
         GridColor       =   12582912
      End
   End
End
Attribute VB_Name = "LOGROS_REPORT_AREA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
If Dir(Ruta & Combo2.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
ret = 0
NAR = FreeFile
Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
MATI50.Rows = 1
Frame1.Caption = ""
fl = Left(Combo2.Text, 1)
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo2.Text Then
        ser = Left(icur.grado, 3)
        RE22 = icur.grado
    End If
Wend
Close #NAR
Y = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    Y = Y + 1
    Get #NAR, Y, mate
    If RTrim(mate.nom) = Combo3.Text Then
        que = mate.num
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
        pio = 1
    End If
Wend
Close #NAR
If pio = 0 Then
    MsgBox "ESTA AREA NO ESTA CREADA PARA ESTE GRADO", 16, "OBSERVACIONES"
    Screen.MousePointer = 0
    Exit Sub
End If
tnumlog = 0
MATI50.Col = 2
'****RUTINA PARA CALCULAR LOGROS TOTALES*****
For h = 1 To 4
If Dir(Ruta & fl & ser & que & h & ".lgr") <> "" Then
    numlog = 0
    FERT = 0
    Open Ruta & fl & ser & que & h & ".lgr" For Random As #NAR Len = Len(logru)
    While Not EOF(NAR)
        FERT = FERT + 1
        Get #NAR, FERT, logru
        If logru.indicador = "L" Then
            numlog = numlog + 1
            tnumlog = tnumlog + 1
        End If
    Wend
    Close #NAR
    If h = 1 Then
        periodo = "PRIMERO (" & numlog & ")"
        S1 = numlog
    End If
    If h = 2 Then
        periodo = "SEGUNDO (" & numlog & ")"
        S2 = numlog
    End If
    If h = 3 Then
        periodo = "TERCERO (" & numlog & ")"
        S3 = numlog
    End If
    If h = 4 Then
        periodo = "CUARTO (" & numlog & ")"
        S4 = numlog
    End If
    MATI50.TextMatrix(0, MATI50.Col) = periodo
End If
MATI50.Col = MATI50.Col + 1
Next h
MATI50.TextMatrix(0, 6) = "TOTAL (" & tnumlog & ")"
        
'****RUTINA PARA ENCONTRAR LOGROS, RECUPERACIONES
'****Y DEFICIENCIAS POR PERIODO
MATI50.Rows = ret
MATI50.Row = 0
For VV = 1 To ret - 1
        tl = 0
        tr = 0
        td = 0
        MATI50.Col = 2
        MATI50.Row = MATI50.Row + 1
        Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
        Get #NAR, VV, alugru
        Close #NAR
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        MATI50.TextMatrix(MATI50.Row, 0) = VV
        MATI50.TextMatrix(MATI50.Row, 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        
                    
        For h = 1 To 4
            If Dir(Ruta & Combo2.Text & que & h & ".obs") <> "" Then
                z = 0
                Open Ruta & Combo2.Text & que & h & ".obs" For Random As #NAR Len = Len(notas)
                While Not EOF(NAR)
                    z = z + 1
                    Get #NAR, z, notas
                    If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                        SL = 0
                        SR = 0
                        SD = 0
                        For I = 1 To 10
                                If notas.area(I) <> 0 Then
                                    NAR = FreeFile
                                    Open Ruta & fl & ser & que & h & ".lgr" For Random As #NAR Len = Len(logru)
                                    Get #NAR, notas.area(I), logru
                                    Close #NAR
                                    NAR = NAR - 1
                                    If logru.indicador = "L" Then
                                        SL = SL + 1
                                        tl = tl + 1
                                    End If
                                    If logru.indicador = "R" Then
                                        SR = SR + 1
                                        tr = tr + 1
                                    End If
                                    If logru.indicador = "D" Then
                                        SD = SD + 1
                                        td = td + 1
                                    End If
                                End If
                        Next I
                        If h = 1 Then LTT = S1
                        If h = 2 Then LTT = S2
                        If h = 3 Then LTT = S3
                        If h = 4 Then LTT = S4
                        If (LTT <> 0) Then
                            PPOR = Format(((SL + SR) / LTT) * 100, "##.##")
                            MATI50.TextMatrix(MATI50.Row, MATI50.Col) = "L=" & SL & ",R=" & SR & ",D=" & SD & " (" & PPOR & "%)"
                        End If
                    End If
                Wend
                Close #NAR
            End If
            MATI50.Col = MATI50.Col + 1
    Next h
    If (tnumlog <> 0) Then
        TPPOR = Format(((tl + tr) / tnumlog) * 100, "##.##")
        MATI50.TextMatrix(MATI50.Row, 6) = "L=" & tl & ",R=" & tr & ",D=" & td & " (" & TPPOR & "%)"
    End If
Next VV
Frame1.Caption = "GRUPO:" & Combo2.Text & "    AREA:" & Combo3.Text
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
If MATI50.Rows = 1 Then
    MsgBox "NO EXISTE INFORMACION PARA IMPRIMIR", 16
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton1, "IMPRIMIR")
If RESP = vbYes Then
    Screen.MousePointer = 11
    Printer.Orientation = 2
    Printer.PaperSize = 5
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    Printer.CurrentY = 1
    Printer.CurrentX = 10
    Printer.Print "REPORTE PORCENTAJE DE LOGROS POR PERIODO"
    Printer.CurrentY = 1.5
    Printer.CurrentX = 1
    Printer.Print ini.nombre
    Printer.CurrentX = 1
    Printer.Print Frame1.Caption;
    Printer.CurrentX = 24
    Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print "CD";
    Printer.CurrentX = 1.5
    Printer.Print "APELLIDOS Y NOMBRES";
    VX = 8.5
    For I = 1 To 5
        Printer.CurrentX = VX
        Printer.Print MATI50.TextMatrix(0, I + 1);
        VX = VX + 4.3
    Next I
    Printer.Print ""
    Printer.Print ""
    For I = 1 To MATI50.Rows - 1
        Printer.CurrentX = 0.5
        Printer.Print I;
        Printer.CurrentX = 1.5
        Printer.Print MATI50.TextMatrix(I, 1);
        VX = 8.5
        For J = 1 To 5
            Printer.CurrentX = VX
            Printer.Print MATI50.TextMatrix(I, J + 1);
            VX = VX + 4.3
        Next J
        Printer.Print ""
    Next I
    Printer.EndDoc
End If
Printer.Orientation = 1
Printer.PaperSize = 1
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
MATI50.Row = 0
MATI50.Col = 0
MATI50.ColWidth(0) = 500
MATI50.Text = "COD"
MATI50.Col = 1
MATI50.ColWidth(1) = 4200
MATI50.Text = "APELLIDOS Y NOMBRES"
MATI50.Col = 2
MATI50.ColWidth(2) = 1750
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "PRIMERO"
MATI50.Col = 3
MATI50.ColWidth(3) = 1750
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "SEGUNDO"
MATI50.Col = 4
MATI50.ColWidth(4) = 1750
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "TERCERO"
MATI50.Col = 5
MATI50.ColWidth(5) = 1750
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "CUARTO"
MATI50.Col = 6
MATI50.ColWidth(6) = 1800
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "TOTAL"
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    Command1.Enabled = True
    Command2.Enabled = True
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
            Combo3.AddItem RTrim(mate.nom)
        End If
    Next I
    Close #NAR
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo2.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Combo2.Text = Combo2.List(0)
    Combo3.Text = Combo3.List(0)
Else
    Command1.Enabled = False
    Command2.Enabled = False
End If
Frame1.Caption = ""
End Sub

