VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form REPORTE_PORCENT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de porcentaje de logros y desempeños"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10695
   Icon            =   "ReportePorcent.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   9000
      Picture         =   "ReportePorcent.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   8415
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7080
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   5415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "MATERIA:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   630
      End
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
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin MSFlexGridLib.MSFlexGrid Mtx_Reporte 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   7
         FixedRows       =   2
         FixedCols       =   2
      End
   End
End
Attribute VB_Name = "REPORTE_PORCENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Lgr_Ttl As Integer, PorcentLogro As Single, PromLogros As Single, SumDesemp As Long, ww As Integer, ValiNota As Boolean
Dim TDesemp1 As Byte, TDesemp2 As Byte, TDesemp3 As Byte, TDesemp4 As Byte, FlagDesemp As Boolean
Dim VeriManual As Boolean, ConfLgr As Byte, PorcentManual(10) As Integer, ContPorcent As Integer
Dim AcumulaPorcent As Byte, NotAcumula As Single, DEF_AcumulaPorcent As Byte, DEF_NotAcumula As Single
Dim OkObs As Boolean, OkDes As Boolean, TtlMatX As Integer, ContTtlMat As Integer

Command2.Enabled = False
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If

Mtx_Reporte.Rows = 2
Mtx_Reporte.TextMatrix(0, 2) = ""
Mtx_Reporte.TextMatrix(0, 3) = ""
Mtx_Reporte.TextMatrix(0, 4) = ""
Mtx_Reporte.TextMatrix(0, 5) = ""
Screen.MousePointer = 11
NAR = FreeFile
TN = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    TN = TN + 1
    Get #NAR, TN, mate
    If RTrim(mate.nom) = Combo2.Text Then
        que = mate.num
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
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = RTrim(Combo1.Text) Then
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
    If RTrim(argra.grado) = RE22 And RTrim(argra.nom_grup) = Combo1.Text And argra.num_area = que Then
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
MsgBox "NO SE HA CREADO LA MATERIA " & Combo2.Text & " PARA ESTE GRUPO O NO LE CORRESPONDE", 64, "ADVERTENCIA"
    Combo2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If

Screen.MousePointer = 11
NAR = FreeFile
Frame1.Caption = Combo1.Text & " (" & Combo2.Text & ")" & " - PROFESOR(A): " & PRO

ret = 0
NAR = FreeFile
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR

For J = 1 To (ret - 1)
    NAR = FreeFile
    Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
    Get #NAR, J, alugru
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, (Val(alugru.num_carnet)), alumno
    Close #NAR
    NAR = NAR - 1
    Mtx_Reporte.Rows = J + 2
    Mtx_Reporte.TextMatrix(J + 1, 0) = J
    Mtx_Reporte.TextMatrix(J + 1, 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    cona = 0
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
    'CP = 0
    SumDesemp = 0
    'ValiNota = False
    DEF_AcumulaPorcent = 0
    DEF_NotAcumula = 0
    For ww = 1 To 5
        If ww <> 5 Then
            OkDes = False
            ValiNota = False
            If Dir(Ruta & Trim(Combo1.Text) & que & ww & ".dsp") <> "" Then
            
                NAR = FreeFile
                VV = 0
                Open Ruta & Trim(Combo1.Text) & que & ww & ".dsp" For Random As #NAR Len = Len(notas_desemp)
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
                Open Ruta & Left(Combo1.Text, 1) & Left(RE22, 3) & que & ww & ".lgr" For Random As #NAR Len = Len(logru)
                
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
                    Open Ruta & Left(Combo1.Text, 1) & Left(RE22, 3) & que & ww & ".lgr" For Random As #NAR Len = Len(logru)
                    Get #NAR, notas_desemp.logro(I), logru
                    
                    If notas_desemp.porcentaje(I) <> 0 Then
                        NAR = FreeFile
                        Open Ruta & Left(Combo1.Text, 1) & Left(RE22, 3) & que & ww & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                        Get #NAR, I, porcent_manual
                        Close #NAR
                        NAR = NAR - 1
                        
                        AcumulaPorcent = AcumulaPorcent + porcent_manual.porcent_logro
                        NotAcumula = NotAcumula + (Val(porcent_manual.porcent_logro) * Val(notas_desemp.porcentaje(I)))
                        'ACUMULADO PARA OBTENER LA DEFENITIVA DE TODOS LOS PERIODOS
                        DEF_AcumulaPorcent = DEF_AcumulaPorcent + porcent_manual.porcent_logro
                        DEF_NotAcumula = DEF_NotAcumula + (Val(porcent_manual.porcent_logro) * Val(notas_desemp.porcentaje(I)))
                    End If
                    Close #NAR
                    NAR = NAR - 1
                        
                Next I
                If AcumulaPorcent <> 0 Then
                    'SI VA PERDIENDO MUESTRA EL VR. EN NEGRILLA
                    If (NotAcumula / AcumulaPorcent) < 70 Then
                        Mtx_Reporte.Col = ww + 1
                        Mtx_Reporte.Row = J + 1
                        Mtx_Reporte.CellFontBold = True
                        Mtx_Reporte.Text = Format(NotAcumula / AcumulaPorcent, "#.00")
                     Else
                        Mtx_Reporte.Col = ww + 1
                        Mtx_Reporte.Row = J + 1
                        Mtx_Reporte.CellFontBold = False
                        Mtx_Reporte.Text = Format(NotAcumula / AcumulaPorcent, "#.00")
                    End If
                    Mtx_Reporte.TextMatrix(0, ww + 1) = AcumulaPorcent & "%"
                Else
                    Mtx_Reporte.TextMatrix(J + 1, ww + 1) = ""
                End If
            End If
        Else
            If DEF_AcumulaPorcent <> 0 Then
                'SI VA PERDIENDO EN LA DEFINITIVA MUESTRA EL VR. EN NEGRILLA
                If (DEF_NotAcumula / DEF_AcumulaPorcent) < 70 Then
                    Mtx_Reporte.Col = ww + 1
                    Mtx_Reporte.Row = J + 1
                    Mtx_Reporte.CellFontBold = True
                    Mtx_Reporte.Text = Format(DEF_NotAcumula / DEF_AcumulaPorcent, "#.00")
                 Else
                    Mtx_Reporte.Col = ww + 1
                    Mtx_Reporte.Row = J + 1
                    Mtx_Reporte.CellFontBold = False
                    Mtx_Reporte.Text = Format(DEF_NotAcumula / DEF_AcumulaPorcent, "#.00")
                End If
            Else
                Mtx_Reporte.TextMatrix(J + 1, ww + 1) = ""
            End If
        End If
        
    Next ww
    Close #NAR
Next J
Close #NAR
Mtx_Reporte.Rows = J + 2
Mtx_Reporte.Col = 1
Mtx_Reporte.Row = Mtx_Reporte.Rows - 1
Mtx_Reporte.CellFontBold = True
Mtx_Reporte.Text = "TOTAL PERDIDA..."
'Mostrar el total de estudiantes perdiendo por periodo
For h = 2 To 6
    TtlMatX = 0
    For J = 2 To Mtx_Reporte.Rows - 2
        If Val(Mtx_Reporte.TextMatrix(J, h)) < 70 And Mtx_Reporte.TextMatrix(J, h) <> "" Then
            TtlMatX = TtlMatX + 1
        End If
    Next J
    If TtlMatX <> 0 Then
        Mtx_Reporte.TextMatrix(J, h) = TtlMatX
    End If
Next h

'Mostrar el porcentaje acumulado
TtlMatX = 0
For h = 2 To 5
    TtlMatX = TtlMatX + Val(Mtx_Reporte.TextMatrix(0, h))
Next h
Mtx_Reporte.TextMatrix(0, 6) = TtlMatX & "%"

Screen.MousePointer = 0
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Dim SaltoColum As Single
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
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
   Printer.CurrentX = 1
   Printer.Print ini.nombre;
   Printer.CurrentX = 16.5
   Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
   Printer.Print ""
   Printer.Font.Size = 8
   Printer.CurrentX = 1
   Printer.Print Frame1.Caption;
   Printer.CurrentY = 3
   For J = 0 To Mtx_Reporte.Rows - 1
        For I = 0 To 6
            If I = 0 Then
                SaltoColum = 1
            End If
            If I = 1 Then
                SaltoColum = 2
            End If
            If I = 2 Then
                SaltoColum = 11
            End If
            If I = 3 Then
                SaltoColum = 13
            End If
            If I = 4 Then
                SaltoColum = 15
            End If
            If I = 5 Then
                SaltoColum = 17
            End If
            If I = 6 Then
                SaltoColum = 19
            End If
            If Val(Mtx_Reporte.TextMatrix(J, I)) < 70 And J > 1 And I > 1 Then
                Printer.FontBold = True
            Else
                Printer.FontBold = False
            End If
            Printer.CurrentX = SaltoColum
            Printer.Print Mtx_Reporte.TextMatrix(J, I);
        Next I
        Printer.Print ""
   Next J
   Printer.EndDoc
   Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Load()
Mtx_Reporte.Row = 1
Mtx_Reporte.ColWidth(0) = 500
Mtx_Reporte.CellForeColor = RGB(255, 255, 255)
Mtx_Reporte.CellBackColor = RGB(0, 0, 150)
Mtx_Reporte.Text = "COD"
Mtx_Reporte.Col = 1
Mtx_Reporte.ColWidth(1) = 4200
Mtx_Reporte.CellForeColor = RGB(255, 255, 255)
Mtx_Reporte.CellBackColor = RGB(0, 0, 150)
Mtx_Reporte.Text = "APELLIDOS Y NOMBRES"
Mtx_Reporte.Col = 2
Mtx_Reporte.ColWidth(2) = 1000
Mtx_Reporte.CellForeColor = RGB(255, 255, 255)
Mtx_Reporte.CellBackColor = RGB(0, 0, 150)
Mtx_Reporte.Text = "PRIMERO"
Mtx_Reporte.Col = 3
Mtx_Reporte.ColWidth(3) = 1000
Mtx_Reporte.CellForeColor = RGB(255, 255, 255)
Mtx_Reporte.CellBackColor = RGB(0, 0, 150)
Mtx_Reporte.Text = "SEGUNDO"
Mtx_Reporte.Col = 4
Mtx_Reporte.ColWidth(4) = 1000
Mtx_Reporte.CellForeColor = RGB(255, 255, 255)
Mtx_Reporte.CellBackColor = RGB(0, 0, 150)
Mtx_Reporte.Text = "TERCERO"
Mtx_Reporte.Col = 5
Mtx_Reporte.ColWidth(5) = 1000
Mtx_Reporte.CellForeColor = RGB(255, 255, 255)
Mtx_Reporte.CellBackColor = RGB(0, 0, 150)
Mtx_Reporte.Text = "CUARTO"
Mtx_Reporte.Col = 6
Mtx_Reporte.ColWidth(6) = 1000
Mtx_Reporte.CellForeColor = RGB(255, 255, 255)
Mtx_Reporte.CellBackColor = RGB(0, 0, 150)
Mtx_Reporte.Text = "DEFINITIVA"

If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") Then
    Command1.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo1.AddItem RTrim(icur.nom)
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
            Combo2.AddItem RTrim(mate.nom)
        End If
    Next I
    Close #NAR
    Combo1.Text = Combo1.List(0)
    Combo2.Text = Combo2.List(0)
Else
    Command1.Enabled = False
End If
Command2.Enabled = False
End Sub
