VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RESUFINA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe final"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "RESUFINA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Copiar"
      Height          =   495
      Left            =   8640
      Picture         =   "RESUFINA.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir grupo"
      Height          =   735
      Left            =   7920
      Picture         =   "RESUFINA.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprime los resumenes finales del grupo seleccionado"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   735
      Left            =   6120
      Picture         =   "RESUFINA.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprime el resumen final del estudiante que se muestra en pantalla"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONSULTAR INFORME"
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   5655
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   315
         Left            =   4680
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO:"
         Height          =   195
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid MATI50 
         Height          =   3975
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         BackColorBkg    =   12632256
         GridColor       =   4194368
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   45
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7800
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5880
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "RESUFINA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ******* FUNCION QUE OBTIENE EL DESEMPEÑO POR PERIODOS Y LA DEFINITIVA (FINAL) *******
Private Function Definitiva(DsPeriodo As Integer) As String
Dim VeriManual As Boolean
Dim w As Integer, CP As Integer, PorcentLogro As Single, PromLogros As Single, SumDesemp As Long, PorcentManual(10) As Integer, ContPorcent As Integer
Definitiva = ""

' VERIFICAR PORCENTAJES DE LOGROS AUTOMATICOS O MANUALES
VeriManual = False
If Dir(Ruta & "conf_logro.edu") <> "" Then
    NAR = FreeFile
    Open Ruta & "conf_logro.edu" For Input As #NAR
    Input #NAR, ConfLgr
    Close #NAR
    NAR = NAR - 1
    If ConfLgr = 1 Then
        VeriManual = True
    End If
End If

'OBTENER ARREGLOS RANGOS DE DESEMPEÑOS Y CONVENSIONES DE DESEMPEÑOS.
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

CP = 0
SumDesemp = 0
Lgr_Ttl = 0

For lwe = 1 To 4
If DsPeriodo = lwe Or DsPeriodo = 5 Then
    If Dir(Ruta & Combo1.Text & argra.num_area & lwe & ".dsp") <> "" Then
        
    
        'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
        NAR = FreeFile
        VV = 0
        Open Ruta & Combo1.Text & argra.num_area & lwe & ".dsp" For Random As #NAR Len = Len(notas_desemp)
        While Not EOF(NAR)
            VV = VV + 1
            Get #NAR, VV, notas_desemp
            If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                GoTo PorcentEncontrado
            End If
        Wend
PorcentEncontrado:
        Close #NAR
        NAR = NAR - 1
        
        'OBTENER TOTAL DE LOGROS
        CROA = 0
        Cont_Lgr = 0
        NAR = FreeFile
        Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
        While Not EOF(NAR)
            CROA = CROA + 1
            Get #NAR, CROA, logru
            If Trim(logru.indicador) = "L" Then
                Cont_Lgr = Cont_Lgr + 1
                Lgr_Ttl = Lgr_Ttl + 1
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
       
        ContPorcent = 0
        If VeriManual = False Then
            ' ******* SI LOS PORCENTAJES SON AUTOMATICOS SE HACE LO SIGUIENTE: *******
            For w = 1 To Cont_Lgr
                'Se verifican los logros mayores al 69% (o rango menor)
                If notas_desemp.porcentaje(w) > confdesemp.rango(3) Then
                    CP = CP + 1
                End If
                If notas_desemp.porcentaje(w) = 0 Then
                    ContPorcent = ContPorcent + 1
                End If
                SumDesemp = SumDesemp + notas_desemp.porcentaje(w)
            Next w
            'Se verifica que tenga grabado porcentajes para obtener el desempeño final
            If ContPorcent = Cont_Lgr Then
                'Definitiva = ""
                'Exit Function
                'Exit For
                GoTo SaltaDesemp
            End If
            PorcentLogro = (100 / Lgr_Ttl) * CP
            If PorcentLogro <> 0 Then
                PorcentLogro = Format(PorcentLogro, "#")
            End If
            PromLogros = SumDesemp / Lgr_Ttl
            If PromLogros >= 0.5 Then
                PromLogros = Format(PromLogros, "#")
            End If
            'REPORTE POR DESEMPEÑOS
            If PorcentLogro <= confdesemp.rango(3) Then
                    ' ***Pierde si el porcentaje total de logros alcanzados es menor o igual al Rango inferior
                Definitiva = Trim(confdesemp.desemp(4))
            Else
                ' ***Si alcanza el porcentaje mínimo de logros, se promedia para obtener el desempeño
                If PromLogros <= confdesemp.rango(3) Then
                    Definitiva = Trim(confdesemp.desemp(3))
                End If
                If (PromLogros > confdesemp.rango(3)) And (PromLogros <= confdesemp.rango(2)) Then
                     Definitiva = Trim(confdesemp.desemp(3))
                End If
                
                If (PromLogros > confdesemp.rango(2)) And (PromLogros <= confdesemp.rango(1)) Then
                    Definitiva = Trim(confdesemp.desemp(2))
                End If
                If PromLogros > confdesemp.rango(1) Then
                    Definitiva = Trim(confdesemp.desemp(1))
                End If
            End If
        
        Else
            If Dir(Ruta & fl & seri & argra.num_area & lwe & ".ptj") = "" Then
                Definitiva = "ERROR"
                Exit Function
            Else
                NAR = FreeFile
                Open Ruta & fl & seri & argra.num_area & lwe & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                For h = 1 To Cont_Lgr
                    Get #NAR, h, porcent_manual
                    PorcentManual(h) = porcent_manual.porcent_logro
                Next h
                Close #NAR
                NAR = NAR - 1
            End If
                  
            ' ******* SI LOS PORCENTAJES SON MANUALES SE HACE LO SIGUIENTE: *******
            For w = 1 To Cont_Lgr
                If notas_desemp.porcentaje(w) = 0 Then
                    ContPorcent = ContPorcent + 1
                End If
                SumDesemp = SumDesemp + (notas_desemp.porcentaje(w) * PorcentManual(w))
            Next w
            'Se verifica que tenga grabado porcentajes para obtener el desempeño final
            If ContPorcent = Cont_Lgr Then
                'Definitiva = ""
                'Exit Function
                'Exit For
                GoTo SaltaDesemp
            End If
            If DsPeriodo <> 5 Then
                PorcentLogro = SumDesemp / 100
                If PorcentLogro <> 0 Then
                    PorcentLogro = Format(PorcentLogro, "#")
                End If
            Else
                PorcentLogro = SumDesemp / (100 * lwe)
                If PorcentLogro <> 0 Then
                    PorcentLogro = Format(PorcentLogro, "#")
                End If
            End If
            'REPORTE POR DESEMPEÑOS
            If PorcentLogro <= confdesemp.rango(3) Then
                Definitiva = Trim(confdesemp.desemp(4))
            End If
            If (PorcentLogro > confdesemp.rango(3)) And (PorcentLogro <= confdesemp.rango(2)) Then
                 Definitiva = Trim(confdesemp.desemp(3))
            End If
            
            If (PorcentLogro > confdesemp.rango(2)) And (PorcentLogro <= confdesemp.rango(1)) Then
                Definitiva = Trim(confdesemp.desemp(2))
            End If
            If PorcentLogro > confdesemp.rango(1) Then
                Definitiva = Trim(confdesemp.desemp(1))
            End If
        End If
    
    End If
'Else
'    Exit For
SaltaDesemp:
End If
Next lwe
End Function


'Private Sub Combo1_Change()
'If Combo1.Text <> Combo1.List(0) Then
'    Combo1.Text = Combo1.List(0)
'End If
'End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Command1_Click()
Dim DsFinal As String, MateriaX As Integer
Label4.Caption = ""
Label5.Caption = ""
If RTrim(Text2.Text) = "" Then
    MsgBox "ESCRIBA EL CÓDIGO DEL ESTUDIANTE", 16, "CONSULTAR INFORME FINAL"
    Text2.SetFocus
    Exit Sub
End If
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "NO EXISTE ESTE GRUPO", 48, "CONSULTAR INFORME FINAL"
    Combo1.SetFocus
    Exit Sub
End If
Screen.MousePointer = 11
MATI50.Rows = 1
Frame1.Caption = ""
ret = 0
NAR = FreeFile
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
If Val(Text2.Text) > ret - 1 Or Val(Text2.Text) < 1 Then
    MsgBox "CODIGO NO EXISTE EN ESTE GRUPO", 64, "CONSULTAR"
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
Else
    Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
    Get #NAR, Val(Text2.Text), alugru
    Close #NAR
End If

Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, Val(alugru.num_carnet), alumno
Close #NAR
Label4.Caption = Combo1.Text
Label5.Caption = Val(alugru.num_carnet)

Frame1.Caption = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & Text2.Text & ")."
' Abrir parámetros de promoción
If (Dir(Ruta & "rangpro.txt") <> "") And (Dir(Ruta & "promovido.txt") <> "") Then
    Open Ruta & "rangpro.txt" For Input As #NAR
    Input #NAR, rus, fis
    Close #NAR
    Open Ruta & "promovido.txt" For Input As #NAR
    Input #NAR, SAPO2, SAPO3, SAPO4
    Close #NAR
End If
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo1.Text Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
    End If
Wend
Close #NAR
seri = Left(RE22, 3)
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
MateriaX = 0
nf = 1
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
        MATI50.Rows = nf + 1
        MATI50.TextMatrix(nf, 0) = RTrim(mate.nom)
        MATI50.TextMatrix(nf, 1) = argra.ih
        For ww = 1 To 5
            DsFinal = Definitiva(ww)
            If DsFinal <> "" And DsFinal <> "ERROR" Then
                MATI50.TextMatrix(nf, ww + 1) = DsFinal
                ' Se verifica el total de materias perdidas (teniendo en cuenta la nota final)
                If (Trim(DsFinal) = Trim(confdesemp.desemp(4))) And ww = 5 Then
                    MateriaX = MateriaX + 1
                End If
            Else
                MATI50.TextMatrix(nf, ww + 1) = ""
            End If
        Next ww
        nf = nf + 1
    End If
Wend
Close #NAR
If MateriaX <= rus Then
    Label3.Caption = RTrim(SAPO2)
End If
If (MateriaX > rus) And (MateriaX <= fis) Then
    Label3.Caption = RTrim(SAPO3)
End If
If MateriaX > fis Then
    Label3.Caption = RTrim(SAPO4)
End If
Screen.MousePointer = 0
End Sub
Private Sub Command2_Click()
If MATI50.Rows = 1 Or Label4.Caption = "" Or Label5.Caption = "" Then
    MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 48, "IMPRIMIR"
    Exit Sub
End If
FORMATO2.Show 1
End Sub

Private Sub Command3_Click()
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "NO EXISTE ESTE GRUPO", 48, "IMPRIMIR INFORME"
    Combo1.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
If RTrim(icur.nom) = Combo1.Text Then
    RE22 = RTrim(icur.grado)
    JOJI = RTrim(icur.jornada)
End If
Wend
Close #NAR
seri = Left(RE22, 3)
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
ret = 0
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
ret = ret + 1
Get #NAR, ret, alugru
Wend
Close #NAR
INT_RESU.Show 1
End Sub

Private Sub Command4_Click()
If MATI50.Rows = 1 Then
MsgBox "NO HAY INFORMACION PARA COPIAR", 48, "COPIAR"
Exit Sub
End If
COPIHIST.Show 1
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra el resumen del informe académico de todo el año."
End Sub

Private Sub Form_Load()
'Dim icur As inforcur
MATI50.Row = 0
MATI50.Col = 0
MATI50.ColWidth(0) = 3100
MATI50.CellFontBold = True
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "             M A T E R I A"
MATI50.Col = 1
MATI50.ColWidth(1) = 400
MATI50.CellFontBold = True
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "I.H."
MATI50.Col = 2
MATI50.ColWidth(2) = 1000
MATI50.CellFontBold = True
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "PRIMERO"
MATI50.Col = 3
MATI50.ColWidth(3) = 1000
MATI50.CellFontBold = True
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "SEGUNDO"
MATI50.Col = 4
MATI50.ColWidth(4) = 1000
MATI50.CellFontBold = True
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "TERCERO"
MATI50.Col = 5
MATI50.ColWidth(5) = 1000
MATI50.CellFontBold = True
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "CUARTO"
MATI50.Col = 6
MATI50.ColWidth(6) = 1000
MATI50.CellFontBold = True
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "FINAL"

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
Text2.MaxLength = 5
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
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
