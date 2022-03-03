VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PORCENTAJE_GENRL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentaje de logros por grupo"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "PORCENTAJE_GENRL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   600
      Left            =   4920
      Picture         =   "PORCENTAJE_GENRL.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir lista"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   630
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
      ForeColor       =   &H00800000&
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9615
      Begin MSFlexGridLib.MSFlexGrid MATI126 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
      End
   End
End
Attribute VB_Name = "PORCENTAJE_GENRL"
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

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
ret = 0
NAR = FreeFile
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
MATI126.Rows = 1
MATI126.Cols = 2
Frame1.Caption = ""
MATI126.Col = 0
MATI126.ColWidth(0) = 500
MATI126.CellForeColor = RGB(255, 255, 255)
MATI126.CellBackColor = RGB(0, 0, 150)
MATI126.Text = "COD"
MATI126.Col = 1
MATI126.ColWidth(1) = 4200
MATI126.CellForeColor = RGB(255, 255, 255)
MATI126.CellBackColor = RGB(0, 0, 150)
MATI126.Text = "APELLIDOS Y NOMBRES"
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If RTrim(argra.nom_grup) = Combo1.Text Then
        NAR = FreeFile
        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
        Get #NAR, argra.num_area, mate
        Close #NAR
        NAR = NAR - 1
        MATI126.Cols = MATI126.Cols + 1
        MATI126.Col = MATI126.Cols - 1
        MATI126.ColWidth(MATI126.Col) = 2800
        MATI126.CellForeColor = RGB(255, 255, 255)
        MATI126.CellBackColor = RGB(0, 0, 150)
        MATI126.Text = RTrim(mate.nom) & " (" & mate.num & ")"
    End If
Wend
Close #NAR
ret = 0
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
For J = 1 To (ret - 1)
    Get #NAR, J, alugru
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, (Val(alugru.num_carnet)), alumno
    Close #NAR
    NAR = NAR - 1
    MATI126.Rows = J + 1
    MATI126.TextMatrix(J, 0) = J
    MATI126.TextMatrix(J, 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    cona = 0
    CX = 1
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Combo1.Text Then
            CX = CX + 1
            'k = 0
            numlog = 0
            tl = 0
            tr = 0
            td = 0
            For ww = 1 To 4
                'CP = 0
                
                If Dir(Ruta & Left(Combo1.Text, 1) & Left(argra.grado, 3) & argra.num_area & ww & ".lgr") <> "" Then
                    rr = 0
                    NAR = FreeFile
                    Open Ruta & Left(Combo1.Text, 1) & Left(argra.grado, 3) & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
                    While Not EOF(NAR)
                        rr = rr + 1
                        Get #NAR, rr, logru
                        If logru.indicador = "L" Then
                            numlog = numlog + 1
                        End If
                    Wend
                    Close #NAR
                    NAR = NAR - 1
                    If Dir(Ruta & RTrim(argra.nom_grup) & (argra.num_area) & ww & ".obs") <> "" Then
                        r = 0
                        NAR = FreeFile
                        Open Ruta & RTrim(argra.nom_grup) & (argra.num_area) & ww & ".obs" For Random As #NAR Len = Len(notas)
                        While Not EOF(NAR)
                            r = r + 1
                            Get #NAR, r, notas
                            If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                                'EXISALU = True
                                NAR = FreeFile
                                Open Ruta & Left(Combo1.Text, 1) & Left(argra.grado, 3) & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
                                For I = 1 To 10
                                    If notas.area(I) <> 0 Then
                                        Get #NAR, notas.area(I), logru
                                            If (logru.indicador = "L") Then tl = tl + 1
                                            If (logru.indicador = "R") Then tr = tr + 1
                                            If (logru.indicador = "D") Then td = td + 1
                                    End If
                                Next I
                                Close #NAR
                                NAR = NAR - 1
                            End If
                        Wend
                        Close #NAR
                        NAR = NAR - 1
                    End If
                End If
            Next ww
            'z = z + k
            'MATI126.TextMatrix(J, CX) = "L=" & tl & ",R=" & tr & ",D=" & td
            If (numlog <> 0) Then
                If (tl + tr <> 0) Then
                    TFPOR = Format(((tl + tr) / numlog) * 100, "##.##")
                    'MATI126.TextMatrix(J, CX) = TFPOR & "%"
                    MATI126.TextMatrix(J, CX) = TFPOR
                End If
            End If
        End If
    Wend
    Close #NAR
    NAR = NAR - 1
    'If EXISALU = True Then
    '    MATI126.TextMatrix(J, CX) = z & "(" & h & ")"
    'End If
Next J
Close #NAR

MATI126.Rows = J + 1
MATI126.TextMatrix(J, 1) = "TOTAL PORCENTAJE..."
'tttt = 0
For hh = 2 To CX
    ttp = 0
    For w = 1 To J - 1
        If MATI126.TextMatrix(w, hh) = "" Then
            MATI126.TextMatrix(w, hh) = 0
        End If
        ttp = ttp + MATI126.TextMatrix(w, hh)
    Next w
    
    tsub = Format(ttp / (J - 1), "##.##")
    'tttt = tttt + tsub
    MATI126.TextMatrix(J, hh) = tsub & "%"
Next hh
'MATI126.Cols = MATI126.Cols + 1
'MATI126.TextMatrix(J, CX + 1) = tttt / CX
If MATI126.Rows > 1 Then
    MATI126.FixedRows = 1
    MATI126.FixedCols = 2
End If
Frame1.Caption = "GRUPO:" & Combo1.Text
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
'Dim ini As inicio
If Frame1.Caption <> "" Then
    RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
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
        Printer.Print "REPORTE FINAL DE LOGROS ALCANZADOS POR PORCENTAJE"
        Printer.CurrentY = 1.5
        Printer.CurrentX = 1
        Printer.Print ini.nombre
        Printer.CurrentX = 1
        Printer.Print Frame1.Caption;
        Printer.CurrentX = 24
        Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = 1
        Printer.Print "CD";
        Printer.CurrentX = 1.5
        Printer.Print "APELLIDOS Y NOMBRES";
        CX = 8
        For I = 2 To (MATI126.Cols - 1)
            Printer.CurrentX = CX
            If I <> (MATI126.Cols - 1) Then
                Printer.Print Left(MATI126.TextMatrix(0, I), 3) & Right(MATI126.TextMatrix(0, I), 4);
            Else
                Printer.Print "TTL";
            End If
            CX = CX + 1.15
        Next I
        Printer.Print ""
        Printer.Print ""
        For I = 1 To (MATI126.Rows - 1)
            Printer.CurrentX = 1
            Printer.Print MATI126.TextMatrix(I, 0);
            Printer.CurrentX = 1.5
            Printer.Print MATI126.TextMatrix(I, 1);
            CX = 8
            For J = 2 To (MATI126.Cols - 1)
                Printer.CurrentX = CX
                Printer.Print MATI126.TextMatrix(I, J);
                CX = CX + 1.15
            Next J
            Printer.Print ""
        Next I
        Printer.EndDoc
        Printer.Orientation = 1
        Printer.PaperSize = 1
        Screen.MousePointer = 0
    End If
Else
    MsgBox "NO EXISTE INFORMACION PARA IMPRIMIR", 64, "IMPRIMIR"
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra la cantidad de los logros pendientes en cada una de las áreas vistas por el grupo."
End Sub

Private Sub Form_Load()
'Dim icur As inforcur
If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") Then
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
