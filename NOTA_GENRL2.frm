VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form NOTA_GENRL2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe general por materias"
   ClientHeight    =   6540
   ClientLeft      =   1755
   ClientTop       =   1545
   ClientWidth     =   9390
   Icon            =   "NOTA_GENRL2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9390
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de reporte"
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   6240
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "Desempeños"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Porcentajes"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "NOTA_GENRL2.frx":0442
      Left            =   7680
      List            =   "NOTA_GENRL2.frx":0455
      TabIndex        =   0
      Text            =   "PRIMERO"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   705
      Left            =   8160
      Picture         =   "NOTA_GENRL2.frx":0483
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir la lista que aparece en pantalla"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   6015
      Begin VB.CheckBox Check1 
         Caption         =   "Acumulado"
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   320
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
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
      TabIndex        =   4
      Top             =   360
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid MATI126 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6720
      TabIndex        =   8
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "NOTA_GENRL2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TituloPrint As String, AcumulaPrint As String

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
Dim Lgr_Ttl As Integer, PorcentLogro As Single, PromLogros As Single, SumDesemp As Integer, ValiNota As Boolean
Dim TDesemp1 As Byte, TDesemp2 As Byte, TDesemp3 As Byte, TDesemp4 As Byte, FlagDesemp As Boolean, ContPorcent As Integer
MATI126.Rows = 0
MATI126.Cols = 0
Frame1.Caption = ""
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
MATI126.Rows = MATI126.Rows + 2
MATI126.Cols = 3
MATI126.Row = 1
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
MATI126.Col = 2
MATI126.ColWidth(2) = 1000
MATI126.CellForeColor = RGB(255, 255, 255)
MATI126.CellBackColor = RGB(0, 0, 150)
MATI126.Text = "No.CARNET"
NAR = FreeFile
Frame1.Caption = Combo1.Text
If RTrim(Combo2.Text) = "PRIMERO" Then
lw = 1
End If
If RTrim(Combo2.Text) = "SEGUNDO" Then
lw = 2
End If
If RTrim(Combo2.Text) = "TERCERO" Then
lw = 3
End If
If RTrim(Combo2.Text) = "CUARTO" Then
lw = 4
End If
If RTrim(Combo2.Text) = "FINAL" Then
lw = 5
End If

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
    MATI126.Rows = J + 2
    MATI126.TextMatrix(J + 1, 0) = J
    MATI126.TextMatrix(J + 1, 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    MATI126.TextMatrix(J + 1, 2) = alumno.n_carnet
    ' SE INICIA EN LA FILA 2, LA PRIMERA FILA MUESTRA EL TOTAL DE LOGROS ACUMULADOS
    CX = 2
    cona = 0
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Combo1.Text Then
            CX = CX + 1
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            NAR = NAR - 1
            MATI126.Row = 1
            'CREA COLUMNAS SINO ESTAN CREADAS
            If MATI126.Cols < CX + 1 Then
                MATI126.Cols = MATI126.Cols + 1
            End If
            MATI126.Col = CX
            MATI126.ColWidth(MATI126.Col) = 2800
            MATI126.CellForeColor = RGB(255, 255, 255)
            MATI126.CellBackColor = RGB(0, 0, 150)
            MATI126.Text = RTrim(mate.nom) & " (" & mate.num & ")"
            NAR = FreeFile
            Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
            For h = 1 To 14
                Get #NAR, h, confdesemp
                If Trim(argra.grado) = Trim(confdesemp.grado) Then
                    Exit For
                End If
            Next h
            Close #NAR
            NAR = NAR - 1
            Lgr_Ttl = 0
            CP = 0
            SumDesemp = 0
            'ValiNota = False
            For ww = 1 To lw
                ValiNota = False
                If Check1.Value = 0 Then
                    If ww <> lw Then
                        GoTo SaltaP
                    End If
                End If
                If Dir(Ruta & RTrim(argra.nom_grup) & (argra.num_area) & ww & ".dsp") <> "" Then
                    'ValiNota = False
                    Cont_Lgr = 0
                    FERT = 0
                    NAR = FreeFile
                    Open Ruta & Left(Combo1.Text, 1) & Left(argra.grado, 3) & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
                    While Not EOF(NAR)
                        FERT = FERT + 1
                        Get #NAR, FERT, logru
                        If Trim(logru.indicador) = "L" Then
                            Cont_Lgr = Cont_Lgr + 1
                            Lgr_Ttl = Lgr_Ttl + 1
                        End If
                    Wend
                    Close #NAR
                    NAR = NAR - 1
                    r = 0
                    NAR = FreeFile
                    Open Ruta & RTrim(argra.nom_grup) & (argra.num_area) & ww & ".dsp" For Random As #NAR Len = Len(notas_desemp)
                    While Not EOF(NAR)
                        r = r + 1
                        Get #NAR, r, notas_desemp
                        If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                            For I = 1 To Cont_Lgr
                                If notas_desemp.porcentaje(I) > confdesemp.rango(3) Then
                                    CP = CP + 1
                                    ValiNota = True
                                End If
                                'Se valida si tiene nota
                                'If notas_desemp.porcentaje(I) > 0 Then
                                '    ValiNota = True
                                'End If
                                SumDesemp = SumDesemp + notas_desemp.porcentaje(I)
                            Next I
                            GoTo LPA
                        End If
                    Wend
LPA:
                    Close #NAR
                    NAR = NAR - 1
                End If
SaltaP:
            Next ww
            'SI LA MATERIA NO TIENE LOGROS NO SE OBTIENEN PORCENTAJES NI DESEMPEÑOS
            If Lgr_Ttl > 0 Then
                'FORMULA PARA OBTENER EL PORCENTAJE
                PorcentLogro = (100 / Lgr_Ttl) * CP
                If PorcentLogro <> 0 Then
                    PorcentLogro = Format(PorcentLogro, "#")
                End If
                PromLogros = SumDesemp / Lgr_Ttl
                If PromLogros <> 0 Then
                    PromLogros = Format(PromLogros, "#")
                End If
                'If PorcentLogro > 0 And ValiNota = True Then
                If ValiNota = True Then
                    If Option1.Value = True Then
                        'REPORTE POR PORCENTAJE DE LOGROS ALCANZADOS
                        MATI126.TextMatrix(J + 1, CX) = "[" & CP & "]" & PorcentLogro & "%"
                        TituloPrint = "INFORME GENERAL POR MATERIAS - PORCENTAJE DE LOGROS ALCANZADOS"
                    Else
                        'REPORTE POR DESEMPEÑOS
                        If PorcentLogro <= confdesemp.rango(3) Then
                                ' ***Pierde si el porcentaje total de logros alcanzados es menor o igual al Rango inferior
                                'MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(4)) & "(" & Format(SumDesemp / Lgr_Ttl, "##") & "%)"
                                MATI126.Row = J + 1
                                MATI126.Col = CX
                                MATI126.CellFontBold = True
                                MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(4))
                        Else
                            ' ***Si alcanza el porcentaje mínimo de logros, se promedia para obtener el desempeño
                            If PromLogros <= confdesemp.rango(3) Then
                                MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(3))
                            End If
                            If (PromLogros > confdesemp.rango(3)) And (PromLogros <= confdesemp.rango(2)) Then
                                 'MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(3)) & "(" & Format(SumDesemp / Lgr_Ttl, "##") & "%)"
                                 MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(3))
                            End If
                            
                            If (PromLogros > confdesemp.rango(2)) And (PromLogros <= confdesemp.rango(1)) Then
                               'MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(2)) & "(" & Format(SumDesemp / Lgr_Ttl, "##") & "%)"
                               MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(2))
                            End If
                            If PromLogros > confdesemp.rango(1) Then
                                'MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(1)) & "(" & Format(SumDesemp / Lgr_Ttl, "##") & "%)"
                                MATI126.TextMatrix(J + 1, CX) = Trim(confdesemp.desemp(1))
                            End If
                        End If
                        TituloPrint = "INFORME GENERAL POR MATERIAS - DESEMPEÑOS POR MATERIA"
                    End If
                    MATI126.Row = 0
                    MATI126.Col = CX
                    MATI126.CellFontBold = True
                    MATI126.TextMatrix(0, CX) = "LOGROS -->  [" & Lgr_Ttl & "]"
                End If
            End If
        End If
    Wend
    Close #NAR
    NAR = NAR - 1
Next J
Close #NAR

'*******OBTENER TOTALES POR LOGROS******
If Option1.Value = True Then
    For ww = 2 To 5
        For h = 2 To MATI126.Rows - 1

        Next h
    Next ww
'*******OBTENER TOTALES POR DESEMPEÑOS******
Else
    MATI126.Rows = MATI126.Rows + 1
    MATI126.Row = MATI126.Rows - 1
    MATI126.Col = 1
    MATI126.CellFontBold = True
    MATI126.TextMatrix(MATI126.Rows - 1, 1) = "Totales por desempeños [" & Trim(confdesemp.desemp(4)) & "," & Trim(confdesemp.desemp(3)) & "," & Trim(confdesemp.desemp(2)) & "," & Trim(confdesemp.desemp(1)) & "] ..."
    For ww = 3 To MATI126.Cols - 1
        TDesemp1 = 0
        TDesemp2 = 0
        TDesemp3 = 0
        TDesemp4 = 0
        FlagDesemp = False
        For h = 2 To MATI126.Rows - 1
            If Trim(confdesemp.desemp(4)) = MATI126.TextMatrix(h, ww) Then
                   TDesemp4 = TDesemp4 + 1
                   FlagDesemp = True
            End If
            If Trim(confdesemp.desemp(3)) = MATI126.TextMatrix(h, ww) Then
                   TDesemp3 = TDesemp3 + 1
                   FlagDesemp = True
            End If
            If Trim(confdesemp.desemp(2)) = MATI126.TextMatrix(h, ww) Then
                   TDesemp2 = TDesemp2 + 1
                   FlagDesemp = True
            End If
            If Trim(confdesemp.desemp(1)) = MATI126.TextMatrix(h, ww) Then
                   TDesemp1 = TDesemp1 + 1
                   FlagDesemp = True
            End If
        Next h
        If FlagDesemp = True Then
            MATI126.Row = MATI126.Rows - 1
            MATI126.Col = ww
            MATI126.CellFontBold = True
            MATI126.TextMatrix(MATI126.Rows - 1, ww) = "[" & TDesemp4 & "," & TDesemp3 & "," & TDesemp2 & "," & TDesemp1 & "]"
        End If
    Next ww
End If


If MATI126.Rows > 2 Then
    MATI126.FixedRows = 2
    MATI126.FixedCols = 3
End If
Screen.MousePointer = 0
Command2.Enabled = True
If Check1.Value = 1 And lw <> 1 Then
    AcumulaPrint = " (ACUMULADO)"
Else
    AcumulaPrint = ""
End If
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
        Printer.Font.Size = 9
        Printer.CurrentY = 0.5
        Printer.CurrentX = 17.2 - ((Len(ini.nombre) / 3.3) / 2)
        Printer.FontBold = True
        Printer.Print ini.nombre
'        Printer.CurrentX = 16.5 - (Len(ini.Rector) / 5.2) / 2
'        Printer.Print ini.Rector
        Printer.Print ""
        Printer.CurrentX = 17.2 - ((Len(TituloPrint) / 4) / 2)
        Printer.Print TituloPrint
        Printer.FontBold = False
        Printer.Print ""
        Printer.Font.Size = 8
        Printer.CurrentX = 1
        Printer.Print "GRUPO: " & Combo1.Text;
        Printer.CurrentX = 6
        Printer.Print "PERIODO: " & Combo2.Text & AcumulaPrint;
        Printer.CurrentX = 24
        Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
        Printer.CurrentY = 3
'        Printer.CurrentX = 1
'        Printer.Print "CD";
'        Printer.CurrentX = 1.5
'        Printer.Print "APELLIDOS Y NOMBRES";
'
        
        CX = 8
        For I = 3 To (MATI126.Cols - 1)
            Printer.CurrentX = CX
            'If I <> (MATI126.Cols - 1) Then
                Printer.Print Trim(Right(MATI126.TextMatrix(0, I), 5));
            'Else
            '    Printer.Print "TTL";
            'End If
            CX = CX + 1.17
        Next I
        Printer.Print ""
        
        Printer.CurrentX = 1
        Printer.Print "CD";
        Printer.CurrentX = 1.5
        Printer.Print "APELLIDOS Y NOMBRES";
        
        
        CX = 8
        For I = 3 To (MATI126.Cols - 1)
            Printer.CurrentX = CX
            'If I <> (MATI126.Cols - 1) Then
                Printer.Print Left(MATI126.TextMatrix(1, I), 3) & Right(MATI126.TextMatrix(1, I), 4);
            'Else
            '    Printer.Print "TTL";
            'End If
            CX = CX + 1.17
        Next I
        Printer.Print ""
        Printer.Print ""
        For I = 2 To (MATI126.Rows - 1)
            Printer.CurrentX = 1
            Printer.Print MATI126.TextMatrix(I, 0);
            Printer.CurrentX = 1.5
            Printer.Print MATI126.TextMatrix(I, 1);
            CX = 8
            For J = 3 To (MATI126.Cols - 1)
                Printer.CurrentX = CX
                Printer.Print MATI126.TextMatrix(I, J);
                CX = CX + 1.17
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
ENTRADA.stb.Panels.Item(1).Text = "Muestra el juicio valorativo de cada estudiante, en cada área."
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
Command2.Enabled = False
'Check1.Value = 1
Option1.Value = True
End Sub
