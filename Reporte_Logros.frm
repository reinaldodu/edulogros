VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Reporte_Logros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Logros perdidos y reaprendizajes"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9660
   Icon            =   "Reporte_Logros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   7920
      Picture         =   "Reporte_Logros.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Reporte"
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   5280
      TabIndex        =   10
      Top             =   5880
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "Logros con Reaprendizaje"
         Height          =   435
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Logros Perdidos"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "Ver Logros"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MATERIA:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Reporte_Logros.frx":0884
      Left            =   7920
      List            =   "Reporte_Logros.frx":0897
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00800000&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9375
      Begin MSFlexGridLib.MSFlexGrid Mt_Reporte 
         Height          =   4815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8493
         _Version        =   393216
         Rows            =   1
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6120
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
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
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "Reporte_Logros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TituloPrint As String
Private Function CortaObs(Observacion As String)
Dim Recorrer As Integer, Cortar() As String, XSuma As Single
Cortar = Split(Observacion, " ")
XSuma = 0
Printer.CurrentX = 0.5
Printer.FontBold = False
For Recorrer = 0 To UBound(Cortar)
    XSuma = XSuma + Printer.TextWidth(Cortar(Recorrer))
    If XSuma <= 16 Then
        Printer.FontSize = 8
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    Else
        XSuma = Printer.TextWidth(Cortar(Recorrer))
        Printer.Print ""
        Printer.FontSize = 8
        Printer.CurrentX = 0.5
        Printer.FontBold = False
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    End If
Next Recorrer
Printer.Print ""
End Function
Private Sub Command1_Click()
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
Unload Ver_Obser
Command2.Enabled = False
Command3.Enabled = False
Mt_Reporte.Rows = 1
Mt_Reporte.Cols = 2
Label4.Caption = ""
Label5.Caption = ""
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
'Verificamos la configuración de porcentajes para el grado
Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
For h = 1 To 14
    Get #NAR, h, confdesemp
    If RE22 = Trim(confdesemp.grado) Then
        Exit For
    End If
Next h
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
MsgBox "NO SE HA CREADO EL AREA " & Combo3.Text & " PARA ESTE GRUPO O NO LE CORRESPONDE", 64, "ADVERTENCIA"
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
ser = Left(RE22, 3)
FERT = 0
Cont_Lgr = 0
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    FERT = FERT + 1
    Get #NAR, FERT, logru
    If Trim(logru.indicador) = "L" Then
        Cont_Lgr = Cont_Lgr + 1
    End If
Wend
Close #NAR
Label4.Caption = Ruta & fl & ser & que & lw & ".lgr"
If Cont_Lgr = 0 Then
    MsgBox "DEBE GRABAR PRIMERO LOGROS DE " & Combo3.Text & " PARA " & Combo2.Text, 64, "Reporte Logros"
    Screen.MousePointer = 0
    Exit Sub
End If
For h = 1 To Cont_Lgr
    Mt_Reporte.Cols = Mt_Reporte.Cols + 1
    Mt_Reporte.ColWidth(Mt_Reporte.Cols - 1) = 800
    Mt_Reporte.TextMatrix(0, Mt_Reporte.Cols - 1) = "Lgr No." & h
    Mt_Reporte.Row = 0
    Mt_Reporte.Col = (Mt_Reporte.Cols - 1)
    Mt_Reporte.CellForeColor = RGB(255, 255, 255)
    Mt_Reporte.CellBackColor = RGB(0, 0, 150)
Next h
If Dir(Ruta & Combo2.Text & que & lw & ".dsp") <> "" Then
    Y = 0
    Open Ruta & Combo2.Text & que & lw & ".dsp" For Random As #NAR Len = Len(notas_desemp)
    While Not EOF(NAR)
        Y = Y + 1
        Get #NAR, Y, notas_desemp
    Wend
    Close #NAR
    For I = 1 To (Y - 1)
        Mt_Reporte.Rows = I + 1
        Mt_Reporte.TextMatrix(I, 0) = I
        Open Ruta & Combo2.Text & que & lw & ".dsp" For Random As #NAR Len = Len(notas_desemp)
        Get #NAR, I, notas_desemp
        Close #NAR
        For J = 1 To (Cont_Lgr)
            If Option1.Value = True Then
                If notas_desemp.porcentaje(J) <= confdesemp.rango(3) And notas_desemp.porcentaje(J) > 0 Then
                    Mt_Reporte.TextMatrix(I, J + 1) = "      X"
                End If
            Else
                If notas_desemp.recuperado(J) = True Then
                        Mt_Reporte.TextMatrix(I, J + 1) = "      X"
                End If
            End If
        Next J
        If RTrim(notas_desemp.num_carnet) = "" Then
            GoTo salbla
        End If
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(notas_desemp.num_carnet)), alumno
        Close #NAR
        Mt_Reporte.TextMatrix(I, 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        'Mt_Reporte.TextMatrix(I, 1) = Right(alumno.n_carnet, 5)
salbla:
    Next I
Else
    MsgBox "DEBE GRABAR PRIMERO LOS DESEMPEÑOS DE " & Combo3.Text & " PARA " & Combo2.Text, 64, "Reporte Logros"
    Screen.MousePointer = 0
    Exit Sub
End If
'*******OBTENER TOTAL DE LOGROS PERDIDOS******
Mt_Reporte.Rows = Mt_Reporte.Rows + 1
Mt_Reporte.Row = Mt_Reporte.Rows - 1
Mt_Reporte.Col = 1
Mt_Reporte.CellFontBold = True
If Option1.Value = True Then
    Mt_Reporte.TextMatrix(Mt_Reporte.Rows - 1, 1) = "TOTAL LOGROS PERDIDOS..."
    TituloPrint = "LOGROS PERDIDOS"
    Label5.Caption = "REPORTE DE LOGROS PERDIDOS"
Else
    Mt_Reporte.TextMatrix(Mt_Reporte.Rows - 1, 1) = "TOTAL LOGROS CON REAPRENDIZAJE..."
    TituloPrint = "LOGROS CON REAPRENDIZAJE"
    Label5.Caption = "REPORTE DE LOGROS CON REAPRENDIZAJE"
End If
For ww = 2 To Mt_Reporte.Cols - 1
    CP = 0
    FlagDesemp = False
    For h = 1 To Mt_Reporte.Rows - 1
        If Trim(Mt_Reporte.TextMatrix(h, ww)) <> "" Then
            CP = CP + 1
            FlagDesemp = True
        End If
    Next h
    If FlagDesemp = True Then
        Mt_Reporte.Row = Mt_Reporte.Rows - 1
        Mt_Reporte.Col = ww
        Mt_Reporte.CellFontBold = True
        Mt_Reporte.TextMatrix(Mt_Reporte.Rows - 1, ww) = CP
    End If
Next ww

Frame1.Caption = "GRUPO: " & Combo2.Text & " - " & " AREA: " & Combo3.Text & " - " & " PROFESOR(A): " & PRO
Command2.Enabled = True
Command3.Enabled = True
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
SWobserv = False
Ver_Obser.Show
End Sub

Private Sub Command3_Click()
Dim XMax As Single
If Val(Mt_Reporte.Rows - 1) = 0 Then
   MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 48, "IMPRIMIR"
   Exit Sub
End If
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
   Printer.FontBold = False
   Printer.CurrentY = 1
   Printer.CurrentX = 6.5
   'Printer.Print "PORCENTAJE DE LOGROS " & Frame1.Caption
   Printer.Print ""
   Printer.CurrentX = 0.5
   Printer.Print ini.nombre;
   Printer.CurrentX = 12
   If lw = 1 Then
    Printer.Print "PERIODO: PRIMERO";
   End If
   If lw = 2 Then
    Printer.Print "PERIODO: SEGUNDO";
   End If
   If lw = 3 Then
    Printer.Print "PERIODO: TERCERO";
   End If
   If lw = 4 Then
    Printer.Print "PERIODO: CUARTO";
   End If
   If lw = 5 Then
    Printer.Print "PERIODO: FINAL";
   End If
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
   Printer.Print TituloPrint
   'Printer.CurrentX = 16.5
   'Printer.Print "PERIODO: " & Combo1.Text
   Val_X = 10.5
   For I = 2 To Mt_Reporte.Cols - 1
        Printer.CurrentX = Val_X
        Printer.Print "L" & I - 1;
        Val_X = Val_X + 1
   Next I
   Printer.Print ""
    
   For I = 1 To (Mt_Reporte.Rows - 1)
        If I = Mt_Reporte.Rows - 1 Then
            Printer.FontBold = True
            Printer.Print ""
            Printer.CurrentX = 1.3
            Printer.Print RTrim(Mt_Reporte.TextMatrix(I, 1));
        Else
            Printer.FontBold = False
            Printer.CurrentX = 0.5
            Printer.Print I;
            Printer.CurrentX = 1.3
            Printer.Print RTrim(Mt_Reporte.TextMatrix(I, 1));
        End If
      Val_X = 10.5
      For J = 2 To Mt_Reporte.Cols - 1
        Printer.CurrentX = Val_X
        Mt_Reporte.Row = I
        Mt_Reporte.Col = J
        Printer.Print Trim(Mt_Reporte.TextMatrix(I, J));
        Val_X = Val_X + 1
      Next J
      Printer.Print ""
   Next I
   'SE IMPRIMEN LOS LOGROS
   Printer.Print ""
   Printer.Print ""
   Printer.CurrentX = 0.5
   Printer.Print "LOGROS EVALUADOS:"
   'Printer.Print ""
   I = 0
   Cont_Lgr = 0
    NAR = FreeFile
    Open Label4.Caption For Random As #NAR Len = Len(logru)
    While Not EOF(NAR)
        I = I + 1
        Get #NAR, I, logru
        If Trim(logru.indicador) = "L" Then
            Cont_Lgr = Cont_Lgr + 1
            Printer.CurrentX = 0.5
            Printer.FontSize = 8
            Printer.FontBold = False
            'Printer.Print Cont_Lgr;
            'Printer.CurrentX = 1
            XMax = Printer.TextWidth(Trim(logru.indicador) & Cont_Lgr & " - " & Trim(logru.observ))
            If XMax > 16 Then
                CortaObs (Trim(logru.indicador) & Cont_Lgr & " - " & Trim(logru.observ))
            Else
                Printer.Print Trim(logru.indicador) & Cont_Lgr & " - " & Trim(logru.observ)
            End If
        End If
    Wend
    Close #NAR
   Printer.EndDoc
   Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Load()
Mt_Reporte.ColWidth(0) = 400
Mt_Reporte.TextMatrix(0, 0) = "CD"
Mt_Reporte.ColWidth(1) = 3800
Mt_Reporte.TextMatrix(0, 1) = "APELLIDOS Y NOMBRES"
If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") Then
    Command1.Enabled = True
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
    Command1.Enabled = False
End If

Combo1 = Combo1.List(0)
Combo2 = Combo2.List(0)
Combo3 = Combo3.List(0)
Command2.Enabled = False
Command3.Enabled = False
Option1.Value = True
End Sub
