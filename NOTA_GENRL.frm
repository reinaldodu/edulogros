VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form NOTA_GENRL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe general por materias"
   ClientHeight    =   6540
   ClientLeft      =   1755
   ClientTop       =   1545
   ClientWidth     =   9390
   Icon            =   "NOTA_GENRL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9390
   Begin MSFlexGridLib.MSFlexGrid MTorden 
      Height          =   615
      Left            =   8400
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "NOTA_GENRL.frx":0442
      Left            =   7680
      List            =   "NOTA_GENRL.frx":0455
      TabIndex        =   0
      Text            =   "PRIMERO"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   705
      Left            =   7800
      Picture         =   "NOTA_GENRL.frx":0483
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
      Width           =   7335
      Begin VB.CheckBox Check1 
         Caption         =   "Acumulado"
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   320
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
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
Attribute VB_Name = "NOTA_GENRL"
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
Dim Lgr_Ttl As Integer, PorcentLogro As Single, PromLogros As Single, SumDesemp As Long, ValiNota As Boolean
Dim TDesemp1 As Byte, TDesemp2 As Byte, TDesemp3 As Byte, TDesemp4 As Byte, FlagDesemp As Boolean
Dim VeriManual As Boolean, ConfLgr As Byte, PorcentManual(10) As Integer, ContPorcent As Integer
Dim AcumulaPorcent As Byte, NotAcumula As Single, DEF_AcumulaPorcent As Byte, DEF_NotAcumula As Single, SUM_NOTAS As Single
Dim OkObs As Boolean, OkDes As Boolean, TtlMatX As Integer, ContTtlMat As Integer

MATI126.Rows = 0
MATI126.Cols = 0
Frame1.Caption = ""
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
MATI126.Rows = MATI126.Rows + 2
MATI126.Cols = 4
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

' *******Verificar si está disponible la configuración manual de porcentajes de logros*******
'VeriManual = False
'If Dir(Ruta & "conf_logro.edu") <> "" Then
'    NAR = FreeFile
'    Open Ruta & "conf_logro.edu" For Input As #NAR
'    Input #NAR, ConfLgr
'    Close #NAR
'    If ConfLgr = 1 Then
'        VeriManual = True
'    End If
'End If

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
    MATI126.Rows = J + 3
    MATI126.TextMatrix(J + 1, 0) = J
    MATI126.TextMatrix(J + 1, 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    MATI126.TextMatrix(J + 1, 2) = alumno.n_carnet
    CX = 2
    cona = 0
    'SUM_NOTAS = 0
    'TtlMatX = 0
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Combo1.Text Then
            'OkDes = False
            CX = CX + 1
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            NAR = NAR - 1
            MATI126.Row = 1
            'CREA COLUMNAS SINO ESTAN CREADAS
            If MATI126.Cols < CX + 2 Then
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
            DEF_AcumulaPorcent = 0
            DEF_NotAcumula = 0
            For ww = 1 To lw
                OkDes = False
                ValiNota = False
                If Check1.Value = 0 Then
                    If ww <> lw Then
                        GoTo SaltaP
                    End If
                End If
                If Dir(Ruta & RTrim(argra.nom_grup) & (argra.num_area) & ww & ".dsp") <> "" Then
                
                    NAR = FreeFile
                    VV = 0
                    Open Ruta & RTrim(argra.nom_grup) & (argra.num_area) & ww & ".dsp" For Random As #NAR Len = Len(notas_desemp)
                    While Not EOF(NAR)
                        VV = VV + 1
                        Get #NAR, VV, notas_desemp
                        If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                            OkDes = True
                            GoTo encontrar2
                        End If
                    Wend
encontrar2:
                    Close #NAR
                    NAR = NAR - 1
                End If
                
               If OkDes = True Then
                
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
    
                    AcumulaPorcent = 0
                    NotAcumula = 0
                    For I = 1 To Cont_Lgr
                        NAR = FreeFile
                        Open Ruta & Left(Combo1.Text, 1) & Left(argra.grado, 3) & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
                        Get #NAR, notas_desemp.logro(I), logru
                        If notas_desemp.porcentaje(I) <> 0 Then
                            NAR = FreeFile
                            Open Ruta & Left(Combo1.Text, 1) & Left(argra.grado, 3) & argra.num_area & ww & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                            Get #NAR, I, porcent_manual
                            Close #NAR
                            NAR = NAR - 1
                            
                            AcumulaPorcent = AcumulaPorcent + porcent_manual.porcent_logro
                            NotAcumula = NotAcumula + (Val(porcent_manual.porcent_logro) * Val(notas_desemp.porcentaje(I)))
                            'ACUMULADO PARA OBTENER LA DEFENITIVA DE TODOS LOS PERIODOS
                            DEF_AcumulaPorcent = DEF_AcumulaPorcent + porcent_manual.porcent_logro
                            DEF_NotAcumula = DEF_NotAcumula + (Val(porcent_manual.porcent_logro) * Val(notas_desemp.porcentaje(I)))
                            'SUM_NOTA = SUM_NOTA + DEF_NotAcumula
                        End If
                        Close #NAR
                        NAR = NAR - 1
                            
                    Next I
                    If Check1.Value = 0 Then
                        If AcumulaPorcent <> 0 Then
                            'SI VA PERDIENDO MUESTRA EL VR. EN NEGRILLA
                            If (NotAcumula / AcumulaPorcent) < 70 Then
                                MATI126.Col = CX
                                MATI126.Row = J + 1
                                MATI126.CellFontBold = True
                                MATI126.Text = Format(NotAcumula / AcumulaPorcent, "#.00")
                                'TtlMatX = TtlMatX + 1
                             Else
                                MATI126.Col = CX
                                MATI126.Row = J + 1
                                MATI126.CellFontBold = False
                                MATI126.Text = Format(NotAcumula / AcumulaPorcent, "#.00")
                            End If
                            'MUESTRA LOS PORCENTAJES DE LAS MATERIAS
                            MATI126.Col = CX
                            MATI126.Row = 0
                            MATI126.CellFontBold = True
                            MATI126.Text = AcumulaPorcent & "%"
                        Else
                            MATI126.TextMatrix(J + 1, CX) = ""
                        End If
                    Else
                    
                        If DEF_AcumulaPorcent <> 0 Then
                            'SI VA PERDIENDO MUESTRA EL VR. EN NEGRILLA
                            If (DEF_NotAcumula / DEF_AcumulaPorcent) < 70 Then
                                MATI126.Col = CX
                                MATI126.Row = J + 1
                                MATI126.CellFontBold = True
                                MATI126.Text = Format(DEF_NotAcumula / DEF_AcumulaPorcent, "#.00")
                                'TtlMatX = TtlMatX + 1
                             Else
                                MATI126.Col = CX
                                MATI126.Row = J + 1
                                MATI126.CellFontBold = False
                                MATI126.Text = Format(DEF_NotAcumula / DEF_AcumulaPorcent, "#.00")
                            End If
                            'MUESTRA LOS PORCENTAJES DE LAS MATERIAS
                            MATI126.Col = CX
                            MATI126.Row = 0
                            MATI126.CellFontBold = True
                            'Verificar el % acumulado mayor
                            If Val(MATI126.Text) < DEF_AcumulaPorcent Then
                                MATI126.Text = DEF_AcumulaPorcent & "%"
                            End If
                        Else
                            MATI126.TextMatrix(J + 1, CX) = ""
                        End If
                    
                    End If
                End If
SaltaP:
            
            Next ww
        End If
    Wend
    Close #NAR
    NAR = NAR - 1
    'Muestra el total de materias perdidas por estudiante
'    MATI126.Col = CX + 1
'    MATI126.Row = 1
'    MATI126.ColWidth(MATI126.Col) = 2800
'    MATI126.CellForeColor = RGB(255, 255, 255)
'    MATI126.CellBackColor = RGB(0, 0, 150)
'    MATI126.Text = "TOTAL"
    'MATI126.TextMatrix(J + 1, CX + 1) = TtlMatX
Next J
Close #NAR

'CONTAR TOTAL DE MATERIAS PERDIDAS POR ESTUDIANTE
MATI126.Col = MATI126.Cols - 1
MATI126.Row = 1
MATI126.ColWidth(MATI126.Col) = 1000
MATI126.CellForeColor = RGB(255, 255, 255)
MATI126.CellBackColor = RGB(0, 0, 150)
MATI126.Text = "TOTAL"
For J = 2 To MATI126.Rows - 2
    TtlMatX = 0
    For I = 3 To MATI126.Cols - 2
        If (Val(MATI126.TextMatrix(J, I)) < 70) And (MATI126.TextMatrix(J, I) <> "") Then
            TtlMatX = TtlMatX + 1
        End If
    Next I
    MATI126.Col = I
    MATI126.Row = J
    MATI126.CellFontBold = True
    MATI126.Text = TtlMatX
Next J

'CONTAR TOTAL DE PERDIDAS POR MATERIA
For I = 3 To MATI126.Cols - 2
    ContTtlMat = 0
    For J = 2 To MATI126.Rows - 2
        If (Val(MATI126.TextMatrix(J, I)) < 70) And (MATI126.TextMatrix(J, I) <> "") Then
            ContTtlMat = ContTtlMat + 1
        End If
    Next J
    MATI126.Col = I
    MATI126.Row = MATI126.Rows - 1
    MATI126.CellFontBold = True
    MATI126.Text = ContTtlMat
Next I
MATI126.Col = 1
MATI126.Row = MATI126.Rows - 1
MATI126.CellFontBold = True
MATI126.Text = "TOTAL PERDIDA POR MATERIA..."

'******PUESTO POR ESTUDIANTE*****
MTorden.Rows = 0
MATI126.Cols = MATI126.Cols + 1
MATI126.Col = MATI126.Cols - 1
MATI126.Row = 1
MATI126.ColWidth(MATI126.Col) = 1000
MATI126.CellForeColor = RGB(255, 255, 255)
MATI126.CellBackColor = RGB(0, 0, 150)
MATI126.Text = "PUESTO"

For J = 2 To MATI126.Rows - 2
    'TtlMatX = 0
    SUM_NOTAS = 0
    For I = 3 To MATI126.Cols - 3
        If (MATI126.TextMatrix(J, I) <> "") Then
            SUM_NOTAS = SUM_NOTAS + MATI126.TextMatrix(J, I)
        End If
    Next I
'    MATI126.Col = I + 1
'    MATI126.Row = J
'    MATI126.CellFontBold = True
'    MATI126.Text = SUM_NOTAS
    MTorden.Rows = J - 1
    MTorden.TextMatrix(J - 2, 0) = J - 1
    MTorden.TextMatrix(J - 2, 1) = SUM_NOTAS
Next J
MTorden.Col = 1
MTorden.Sort = 4
For TT = 0 To MTorden.Rows - 1
    MATI126.TextMatrix(MTorden.TextMatrix(TT, 0) + 1, I + 1) = TT + 1
Next TT

Screen.MousePointer = 0
Command2.Enabled = True
If MATI126.Cols > 3 Then
    MATI126.FixedCols = 3
End If
If MATI126.Rows > 2 Then
    MATI126.FixedRows = 2
End If
If Check1.Value = 1 Then
    Frame1.Caption = Frame1.Caption & " (ACUMULADO)"
End If
End Sub

Private Sub Command2_Click()
'Dim ini As inicio
If Frame1.Caption <> "" Then

    Impr_NotaGeneral.Show 1

Else
    MsgBox "NO EXISTE INFORMACION PARA IMPRIMIR", 64, "IMPRIMIR"
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra la nota de cada estudiante, en cada materia."
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
'Option1.Value = True
End Sub
