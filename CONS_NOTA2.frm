VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cons_nota2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta e impresión de boletines - Mitad de periodo"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10365
   Icon            =   "CONS_NOTA2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir Grupo"
      Height          =   855
      Left            =   8640
      Picture         =   "CONS_NOTA2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Impresión de boletines por grupo"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   7080
      Picture         =   "CONS_NOTA2.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Impresión del boletín de acuerdo al código ingresado"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "CONS_NOTA2.frx":0EEE
      Left            =   8760
      List            =   "CONS_NOTA2.frx":0F01
      TabIndex        =   0
      Text            =   "PRIMERO"
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONSULTAR BOLETIN"
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   6375
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   4920
         TabIndex        =   3
         Top             =   480
         Width           =   1035
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3720
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO:"
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
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
      Height          =   4455
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid MATI20 
         Height          =   3855
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   12632256
         GridColor       =   12582912
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3240
         TabIndex        =   12
         Top             =   4680
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "IMPRIMIR BOLETIN"
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   6960
      TabIndex        =   19
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4320
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3120
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1560
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7680
      TabIndex        =   9
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "cons_nota2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OkObs As Boolean, OkDes As Boolean
Private Function CortaObs(Observacion As String)
Dim Recorrer As Integer, Cortar() As String, XSuma As Single
Cortar = Split(Observacion, " ")
XSuma = 0
Printer.CurrentX = 0.5
For Recorrer = 0 To UBound(Cortar)
    XSuma = XSuma + Printer.TextWidth(Cortar(Recorrer))
    If XSuma <= 13.3 Then
        Printer.FontSize = 8
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    Else
        XSuma = Printer.TextWidth(Cortar(Recorrer))
        Printer.Print ""
        Call SaltaLinea
        Printer.FontSize = 8
        Printer.CurrentX = 0.5
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    End If
Next Recorrer
Printer.Print ""
End Function

Private Function SaltaLinea()
If (Printer.CurrentY > 26) Then
   Printer.Line (15.7, 5)-(15.7, Printer.CurrentY)
   Printer.Line (17.9, 5)-(17.9, Printer.CurrentY)
   Printer.NewPage
   Call Encabezado
End If
End Function

Private Function Encabezado()
    Printer.ScaleMode = 7
    'Printer.CurrentY = 0.5
    'Printer.Font.Size = 14
    'Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 2)
    'Printer.FontBold = True
    'Printer.Print ini.nombre
    'Printer.CurrentX = 7.4
    'Printer.Print "INFORME ACADÉMICO"
    'Printer.FontBold = False
    'Printer.Print ""
    
    Printer.Font.Size = 10
    Printer.CurrentY = 2.2
    'Printer.CurrentX = 5.5
    'Printer.Print Format(vini.VPeriodo, ">") & ": "  & Combo3.Text
    ' ******** IMPRIME ENCABEZADO ADICIONAL DEL REPORTE DE MITAD DE TRIMESTRE *********
    Printer.CurrentX = (22 - Printer.TextWidth(ConfTexto)) / 2
    Printer.Print ConfTexto
    
    Printer.CurrentY = 3
    Printer.Font.Size = 10
    Printer.CurrentX = 0.5
    Printer.Print Format(vini.VEstudiante, ">") & ": " & Frame1.Caption;
    Printer.CurrentX = 16.5
    Printer.Print Format(vini.VFecha, ">") & ": " & Format(Format(Date, "mmm/dd/yyyy"), ">")
    Printer.CurrentX = 0.5
    Printer.Print Format(vini.VGrupo, ">") & ": " & Combo4.Text
    
    'Printer.Print ""
    ''Printer.CurrentY = Printer.CurrentY + 1
    'Printer.Font.Size = 12
    'Printer.FontBold = True
    'Printer.CurrentX = 0.5
    'Printer.Print "MATERIAS";
    'Printer.Font.Size = 8
    'Printer.CurrentX = 16
    'Printer.CurrentY = Printer.CurrentY + 0.2
    'Printer.Print "PORCENTAJE";
    'Printer.CurrentX = 18
    'Printer.Print "DESEMPEÑO"
    'Printer.FontBold = False
    ''Printer.Font.Size = 10
    'Printer.Line (0.5, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)
    'Printer.Line (0.5, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)
    Printer.CurrentY = 5.5
End Function

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
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Dir(Ruta & Combo4.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
MATI20.Rows = 1
Screen.MousePointer = 11
If RTrim(Text2.Text) = "" Then
    MsgBox "ESCRIBA EL CODIGO DEL ESTUDIANTE", 16, "CONSULTAR OBSERVACIONES"
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
Frame1.Caption = ""
ret = 0
NAR = FreeFile
Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
    'If Val(Text2.Text) = ret Then
    '    Close #NAR
    '    GoTo MIJO
    'End If
Wend
Close #NAR
If Val(Text2.Text) > ret - 1 Or Val(Text2.Text) < 1 Then
    MsgBox "CODIGO NO EXISTE EN ESTE GRUPO", 64, "CONSULTAR"
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
Else
    Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
    Get #NAR, Val(Text2.Text), alugru
    Close #NAR
End If

'MIJO:
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, Val(alugru.num_carnet), alumno
Close #NAR
Frame1.Caption = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & Text2.Text & ")."
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo4.Text Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
    End If
Wend
Close #NAR
Label9.Caption = RE22
'Label10.Caption = RTrim(Combo3.Text)
'Label11.Caption = Combo4.Text
'Label17.Caption = alugru.num_carnet
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
If Combo3.Text = "PRIMERO" Then
    lwe = 1
End If
If Combo3.Text = "SEGUNDO" Then
    lwe = 2
End If
If Combo3.Text = "TERCERO" Then
    lwe = 3
End If
If Combo3.Text = "CUARTO" Then
    lwe = 4
End If
If Combo3.Text = "FINAL" Then
    lwe = 5
End If
cona = 0
'h = 1
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If RTrim(argra.nom_grup) = Combo4.Text Then
        OkObs = False
        OkDes = False
        If Dir(Ruta & Combo4.Text & argra.num_area & lwe & ".obs") <> "" Then
            NAR = FreeFile
            Y = 0
            Open Ruta & Combo4.Text & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
            While Not EOF(NAR)
                Y = Y + 1
                Get #NAR, Y, notas
                If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                    OkObs = True
                    NAR = FreeFile
                    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                    Get #NAR, argra.num_area, mate
                    Close #NAR
                    NAR = NAR - 1
                    MATI20.Rows = MATI20.Rows + 1
                    MATI20.Col = 0
                    MATI20.Row = MATI20.Rows - 1
                    MATI20.CellFontBold = True
                    MATI20.CellForeColor = RGB(0, 0, 255)
                    MATI20.TextMatrix(MATI20.Rows - 1, 0) = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:" & notas.FA & ")"
                    GoTo encontrar
                End If
            Wend
encontrar:
            Close #NAR
            NAR = NAR - 1
        End If
        If Dir(Ruta & Combo4.Text & argra.num_area & lwe & ".dsp") <> "" Then
            NAR = FreeFile
            VV = 0
            Open Ruta & Combo4.Text & argra.num_area & lwe & ".dsp" For Random As #NAR Len = Len(notas_desemp)
            While Not EOF(NAR)
                VV = VV + 1
                Get #NAR, VV, notas_desemp
                If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                    OkDes = True
                    If OkObs = False Then
                        NAR = FreeFile
                        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                        Get #NAR, argra.num_area, mate
                        Close #NAR
                        NAR = NAR - 1
                        MATI20.Rows = MATI20.Rows + 1
                        MATI20.Col = 0
                        MATI20.Row = MATI20.Rows - 1
                        MATI20.CellFontBold = True
                        MATI20.CellForeColor = RGB(0, 0, 255)
                        MATI20.TextMatrix(MATI20.Rows - 1, 0) = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:0)"
                    End If
                    GoTo encontrar2
                End If
            Wend
encontrar2:
            Close #NAR
            NAR = NAR - 1
        End If
        If OkDes = True Then
            CROA = 0
            Cont_Lgr = 0
            NAR = FreeFile
            Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
            While Not EOF(NAR)
                CROA = CROA + 1
                Get #NAR, CROA, logru
                If Trim(logru.indicador) = "L" Then
                    Cont_Lgr = Cont_Lgr + 1
                    
                End If
            Wend
            Close #NAR
            NAR = NAR - 1
            
            NAR = FreeFile
            Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
            For h = 1 To 14
                'Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
                Get #NAR, h, confdesemp
                If Trim(argra.grado) = Trim(confdesemp.grado) Then
                    Exit For
                End If
            Next h
            Close #NAR
            NAR = NAR - 1
            
            For I = 1 To Cont_Lgr
                If notas_desemp.porcentaje(I) <> 0 And notas_desemp.porcentaje(I) <= confdesemp.rango(3) Then
                    NAR = FreeFile
                    Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                    Get #NAR, notas_desemp.logro(I), logru
                    Close #NAR
                    NAR = NAR - 1
                    'If notas_desemp.porcentaje(I) <> 0 And notas_desemp.porcentaje(I) <= confdesemp.rango(3) Then
                        MATI20.Rows = MATI20.Rows + 1
                        MATI20.TextMatrix(MATI20.Rows - 1, 0) = Trim(logru.indicador) & " - " & Trim(logru.observ)
                        MATI20.TextMatrix(MATI20.Rows - 1, 1) = notas_desemp.porcentaje(I) & "%"
                    'End If
                    
                    'NAR = FreeFile
'                    Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
'                    For h = 1 To 14
'                        'Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
'                        Get #NAR, h, confdesemp
'                        If Trim(argra.grado) = Trim(confdesemp.grado) Then
'                            Exit For
'                        End If
'                    Next h
'                    Close #NAR
                    'NAR = NAR - 1
                    'VALIDA RANGOS POR DESEMPEÑO
                     If notas_desemp.porcentaje(I) <> 0 Then
                        If notas_desemp.porcentaje(I) <= confdesemp.rango(3) Then
                            If notas_desemp.recuperado(I) = False Then
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = confdesemp.desemp(4)
                            Else
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = confdesemp.recupera(4)
                            End If
                        End If
                        If (notas_desemp.porcentaje(I) > confdesemp.rango(3)) And (notas_desemp.porcentaje(I) <= confdesemp.rango(2)) Then
                            If notas_desemp.recuperado(I) = False Then
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = confdesemp.desemp(3)
                            Else
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = confdesemp.recupera(3)
                            End If
                        End If
                        
                        If (notas_desemp.porcentaje(I) > confdesemp.rango(2)) And (notas_desemp.porcentaje(I) <= confdesemp.rango(1)) Then
                           If notas_desemp.recuperado(I) = False Then
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = confdesemp.desemp(2)
                            Else
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = confdesemp.recupera(2)
                            End If
                        End If
                        If notas_desemp.porcentaje(I) > confdesemp.rango(1) Then
                            If notas_desemp.recuperado(I) = False Then
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = confdesemp.desemp(1)
                            Else
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = confdesemp.recupera(1)
                            End If
                        End If
                    End If
                End If
            Next I
        End If
        
        
        
        If OkObs = True Then
            For I = 1 To 10
                If notas.area(I) <> 0 Then
                    NAR = FreeFile
                    Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                    Get #NAR, notas.area(I), logru
                    If Trim(logru.indicador) <> "L" Then
                        MATI20.Rows = MATI20.Rows + 1
                        MATI20.TextMatrix(MATI20.Rows - 1, 0) = Trim(logru.indicador) & " - " & Trim(logru.observ)
                    End If
                    Close #NAR
                    NAR = NAR - 1
                End If
            Next I
        End If
        If OkDes = True Or OkObs = True Then
            MATI20.Rows = MATI20.Rows + 1
        End If
              
    End If
Wend
Close #NAR
Screen.MousePointer = 0

End Sub

Private Sub Command2_Click()
Dim ValiSalto As Boolean, XMax As Single, GuardaY As Single

If Dir(Ruta & Combo4.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
'MATI20.Rows = 1
'Screen.MousePointer = 11
If RTrim(Text2.Text) = "" Then
    MsgBox "ESCRIBA EL CODIGO DEL ESTUDIANTE", 16, "CONSULTAR OBSERVACIONES"
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
Frame1.Caption = ""
ret = 0
NAR = FreeFile
Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
    'If Val(Text2.Text) = ret Then
    '    Close #NAR
    '    GoTo MIJO
    'End If
Wend
Close #NAR
If Val(Text2.Text) > ret - 1 Or Val(Text2.Text) < 1 Then
    MsgBox "CODIGO NO EXISTE EN ESTE GRUPO", 64, "CONSULTAR"
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
Else
    Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
    Get #NAR, Val(Text2.Text), alugru
    Close #NAR
End If

'MIJO:
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, Val(alugru.num_carnet), alumno
Close #NAR
Frame1.Caption = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & Text2.Text & ")."

Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo4.Text Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
        SP = RTrim(icur.director)
    End If
Wend
Close #NAR
Label9.Caption = RE22
'Label10.Caption = RTrim(Combo3.Text)
'Label11.Caption = Combo4.Text
'Label17.Caption = alugru.num_carnet
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
If Combo3.Text = "PRIMERO" Then
    lwe = 1
End If
If Combo3.Text = "SEGUNDO" Then
    lwe = 2
End If
If Combo3.Text = "TERCERO" Then
    lwe = 3
End If
If Combo3.Text = "CUARTO" Then
    lwe = 4
End If
If Combo3.Text = "FINAL" Then
    lwe = 5
End If
cona = 0
'h = 1

RESP = MsgBox("DESEA IMPRIMIR EL REPORTE DE " & Frame1.Caption & "?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR REPORTE")
If RESP = vbYes Then
    Screen.MousePointer = 11
    'L = 0

    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
    Close #NAR
    Call Encabezado

Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If RTrim(argra.nom_grup) = Combo4.Text Then
        OkObs = False
        OkDes = False
        If Dir(Ruta & Combo4.Text & argra.num_area & lwe & ".obs") <> "" Then
            NAR = FreeFile
            Y = 0
            Open Ruta & Combo4.Text & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
            While Not EOF(NAR)
                Y = Y + 1
                Get #NAR, Y, notas
                'If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                If (Val(notas.num_carnet) = Val(alugru.num_carnet)) Then
                    For z = 1 To 10
                         If notas.area(z) <> 0 Then
                             OkObs = True
                         End If
                     Next z
                    If OkObs = True Then
                        'OkObs = True
                        NAR = FreeFile
                        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                        Get #NAR, argra.num_area, mate
                        Close #NAR
                        NAR = NAR - 1
                        'MATI20.Rows = MATI20.Rows + 1
                        'Printer.Print ""
                        'MATI20.Col = 0
                        'MATI20.Row = MATI20.Rows - 1
                        'MATI20.CellFontBold = True
                        'MATI20.CellForeColor = RGB(0, 0, 255)
                        Call SaltaLinea
                        Printer.CurrentX = 0.5
                        Printer.FontSize = 10
                        Printer.FontBold = True
                        Printer.Print RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:" & notas.FA & ")"
                        Printer.FontBold = False
                        Printer.FontSize = 8
                        GoTo encontrar
                    End If
                End If
            Wend
encontrar:
            Close #NAR
            NAR = NAR - 1
        End If
        If Dir(Ruta & Combo4.Text & argra.num_area & lwe & ".dsp") <> "" Then
            NAR = FreeFile
            VV = 0
            Open Ruta & Combo4.Text & argra.num_area & lwe & ".dsp" For Random As #NAR Len = Len(notas_desemp)
            While Not EOF(NAR)
                VV = VV + 1
                Get #NAR, VV, notas_desemp
                If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                    For z = 1 To 10
                         If notas_desemp.porcentaje(z) <> 0 And notas_desemp.porcentaje(z) <= confdesemp.rango(3) Then
                             OkDes = True
                         End If
                     Next z
                    'OkDes = True
                    If OkObs = False And OkDes = True Then
                        NAR = FreeFile
                        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                        Get #NAR, argra.num_area, mate
                        Close #NAR
                        NAR = NAR - 1
                        'Printer.Print ""
                        'MATI20.Rows = MATI20.Rows + 1
                        'MATI20.Col = 0
                        'MATI20.Row = MATI20.Rows - 1
                        Call SaltaLinea
                        Printer.CurrentX = 0.5
                        Printer.FontSize = 10
                        Printer.FontBold = True
                        'MATI20.CellFontBold = True
                        'MATI20.CellForeColor = RGB(0, 0, 255)
                        Printer.Print RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:0)"
                        Printer.FontBold = False
                        Printer.FontSize = 8
                        'Call SaltaLinea
                    End If
                    GoTo encontrar2
                End If
            Wend
encontrar2:
            Close #NAR
            NAR = NAR - 1
        End If
        If OkDes = True Then
            CROA = 0
            Cont_Lgr = 0
            NAR = FreeFile
            Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
            While Not EOF(NAR)
                CROA = CROA + 1
                Get #NAR, CROA, logru
                If Trim(logru.indicador) = "L" Then
                    Cont_Lgr = Cont_Lgr + 1

                End If
            Wend
            Close #NAR
            NAR = NAR - 1
            
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
            
            For I = 1 To Cont_Lgr
                
                'VALIDA RANGOS POR DESEMPEÑO
                If notas_desemp.porcentaje(I) <> 0 And notas_desemp.porcentaje(I) <= confdesemp.rango(3) Then
                    'Printer.CurrentX = 16.7
                    Call SaltaLinea
                    Printer.FontSize = 8
                    Printer.CurrentX = 18.8
                    If notas_desemp.porcentaje(I) <= confdesemp.rango(3) Then
                        If notas_desemp.recuperado(I) = False Then
                            Printer.Print confdesemp.desemp(4);
                        Else
                            Printer.Print confdesemp.recupera(4);
                        End If
                    End If
                    If (notas_desemp.porcentaje(I) > confdesemp.rango(3)) And (notas_desemp.porcentaje(I) <= confdesemp.rango(2)) Then
                        If notas_desemp.recuperado(I) = False Then
                            Printer.Print confdesemp.desemp(3);
                        Else
                            Printer.Print confdesemp.recupera(3);
                        End If
                    End If

                    If (notas_desemp.porcentaje(I) > confdesemp.rango(2)) And (notas_desemp.porcentaje(I) <= confdesemp.rango(1)) Then
                       If notas_desemp.recuperado(I) = False Then
                            Printer.Print confdesemp.desemp(2);
                        Else
                            Printer.Print confdesemp.recupera(2);
                        End If
                    End If
                    If notas_desemp.porcentaje(I) > confdesemp.rango(1) Then
                        If notas_desemp.recuperado(I) = False Then
                            Printer.Print confdesemp.desemp(1);
                        Else
                            Printer.Print confdesemp.recupera(1);
                        End If
                    End If
                'End If
                NAR = FreeFile
                Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                Get #NAR, notas_desemp.logro(I), logru
                If notas_desemp.porcentaje(I) <> 0 Then
                    Printer.CurrentX = 16.4
                    Printer.Print notas_desemp.porcentaje(I) & "%";
                    '*** SE VERIFICA EL TAMAÑO DEL LOGRO PARA EL SALTO DE LINEA ***
                    XMax = Printer.TextWidth(Trim(logru.indicador) & " - " & Trim(logru.observ))
                    If XMax > 13.3 Then
                        CortaObs (Trim(logru.indicador) & " - " & Trim(logru.observ))
                    Else
                        Call SaltaLinea
                        Printer.CurrentX = 0.5
                        Printer.FontSize = 8
                        Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                    End If
                End If
                Close #NAR
                NAR = NAR - 1
            End If
            Next I
        End If

        If OkObs = True Then
            'Printer.Print ""
            'ValiSalto = False
            For I = 1 To 10
                If notas.area(I) <> 0 Then
                    NAR = FreeFile
                    Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                    Get #NAR, notas.area(I), logru
                    If Trim(logru.indicador) <> "L" And Trim(logru.indicador) <> "O" And Trim(logru.indicador) <> "S" Then
                        'MATI20.Rows = MATI20.Rows + 1
                        'Printer.CurrentX = 0.5
                        'Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                        '*** SE VERIFICA EL TAMAÑO DEL LOGRO PARA EL SALTO DE LINEA ***
                        XMax = Printer.TextWidth(Trim(logru.indicador) & " - " & Trim(logru.observ))
                        If XMax > 13.3 Then
                            CortaObs (Trim(logru.indicador) & " - " & Trim(logru.observ))
                        Else
                            Call SaltaLinea
                            Printer.CurrentX = 0.5
                            Printer.FontSize = 8
                            Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                        End If
                        ValiSalto = True
                    End If
                    Close #NAR
                    NAR = NAR - 1
                End If
            Next I
'            If ValiSalto = True Then
'                Printer.Print ""
'            End If
            'VALIDAR SI EXISTEN "O" O "S" PARA INFORME DE DESARROLLO
            ValiSalto = False
            For I = 1 To 10
                If notas.area(I) <> 0 Then
                    NAR = FreeFile
                    Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                    Get #NAR, notas.area(I), logru
                    If Trim(logru.indicador) = "O" Or Trim(logru.indicador) = "S" Then
                        'MATI20.Rows = MATI20.Rows + 1
                        'Printer.CurrentX = 0.5
                        'Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                        ValiSalto = True
                    End If
                    Close #NAR
                    NAR = NAR - 1
                End If
            Next I

            If ValiSalto = True Then
                Printer.Print ""
                Call SaltaLinea
                Printer.FontUnderline = True
                Printer.CurrentX = 0.5
                Printer.Print "INFORME DE DESARROLLO:"
                Printer.FontUnderline = False

                'Printer.FontBold = False
                For I = 1 To 10
                    If notas.area(I) <> 0 Then
                        NAR = FreeFile
                        Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                        Get #NAR, notas.area(I), logru
                        If Trim(logru.indicador) = "O" Or Trim(logru.indicador) = "S" Then
                            'MATI20.Rows = MATI20.Rows + 1
                            'Printer.CurrentX = 0.5
                            'Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                            '*** SE VERIFICA EL TAMAÑO DEL LOGRO PARA EL SALTO DE LINEA ***
                            XMax = Printer.TextWidth(Trim(logru.indicador) & " - " & Trim(logru.observ))
                            If XMax > 13.3 Then
                                CortaObs (Trim(logru.indicador) & " - " & Trim(logru.observ))
                            Else
                                Call SaltaLinea
                                Printer.CurrentX = 0.5
                                Printer.FontSize = 8
                                Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                            End If
                            'ValiSalto = True
                        End If
                        Close #NAR
                        NAR = NAR - 1
                    End If
                Next I

                'Printer.Print ""
            End If
        End If
        If OkDes = True Or OkObs = True Then
            Printer.Print ""
            'Printer.Line (0.5, Printer.CurrentY)-(20, Printer.CurrentY)
        End If
    Printer.Line (0.5, Printer.CurrentY)-(20, Printer.CurrentY)
    End If
Wend
Close #NAR
' ********* IMPRIMIR PIE *********
Printer.Line (15.7, 5)-(15.7, Printer.CurrentY)
Printer.Line (17.9, 5)-(17.9, Printer.CurrentY)
' VERIFICAR TAMAÑO NECESARIO PARA IMPRIMIR PIE
If Printer.CurrentY > 20.5 Then
'   Printer.Line (16, 5.5)-(16, Printer.CurrentY)
'   Printer.Line (18, 5.5)-(18, Printer.CurrentY)
   Printer.NewPage
   Call Encabezado
End If
Printer.Print ""
Printer.Print ""
Open Ruta & "leyenda.edu" For Input As #NAR
Input #NAR, leye.ly1, leye.ly2, leye.ly3, leye.ly4, leye.ly5, leye.ly6, leye.ly7, leye.ly8
Close #NAR
Printer.Font.Size = 9
Printer.CurrentX = 0.5
Printer.Print leye.ly1
Printer.CurrentX = 0.5
Printer.Print leye.ly2
Printer.CurrentX = 0.5
Printer.Print leye.ly3
Printer.CurrentX = 0.5
Printer.Print leye.ly4
Printer.CurrentX = 0.5
Printer.Print leye.ly5
Printer.CurrentX = 0.5
Printer.Print leye.ly6
'Printer.Print ""
GuardaY = Printer.CurrentY
'IMPRIMIR CUADRO EXPLICATIVO DE DESEMPEÑOS
Printer.FontSize = 8
Printer.FontBold = True
Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
Printer.CurrentX = 1.7
Printer.Print "DESEMPEÑOS";
Printer.CurrentX = 5.3
Printer.Print "PORCENTAJES";
Printer.CurrentX = 8
Printer.Print "REAPRENDIZAJE"
Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
Printer.FontBold = False
'Printer.CurrentX = 0.7
'Printer.Print confdesemp.desemp(1) & "(Desempeño Superior)";
'Printer.CurrentX = 5.5
'Printer.Print confdesemp.rango(1) + 1 & "% - 100%";
'Printer.CurrentX = 8.8
'Printer.Print confdesemp.recupera(1)
'Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
'Printer.CurrentX = 0.7
'Printer.Print confdesemp.desemp(2) & "(Desempeño Alto)";
'Printer.CurrentX = 5.5
'Printer.Print confdesemp.rango(2) + 1 & "% - " & confdesemp.rango(1) & "%";
'Printer.CurrentX = 8.8
'Printer.Print confdesemp.recupera(2)
'Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
'Printer.CurrentX = 0.7
'Printer.Print confdesemp.desemp(3) & "(Desempeño Básico)";
'Printer.CurrentX = 5.5
'Printer.Print confdesemp.rango(3) + 1 & "% - " & confdesemp.rango(2) & "%";
'Printer.CurrentX = 8.8
'Printer.Print confdesemp.recupera(3)
'Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
Printer.CurrentX = 0.7
Printer.Print confdesemp.desemp(4);
If Trim(confdesemp.desemp(4)) = "*LEP" Then
    Printer.Print " (Logro en Proceso)";
Else
    Printer.Print " (Desempeño Bajo)";
End If
Printer.CurrentX = 5.5
Printer.Print "  0% - " & confdesemp.rango(3) & "%";
Printer.CurrentX = 8.8
Printer.Print confdesemp.recupera(4)
Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
'Trazar líneas verticales
Printer.Line (0.5, GuardaY)-(0.5, Printer.CurrentY)
Printer.Line (5, GuardaY)-(5, Printer.CurrentY)
Printer.Line (7.8, GuardaY)-(7.8, Printer.CurrentY)
Printer.Line (10.6, GuardaY)-(10.6, Printer.CurrentY)
Printer.Print
Printer.CurrentX = 0.5
Printer.Print "LEP = Logro en proceso, que a la fecha está reportado como perdido."
Printer.FontSize = 9
Printer.CurrentX = 0.5
Printer.Print leye.ly7
Printer.CurrentX = 0.5
Printer.Print leye.ly8
'Printer.CurrentX = 0.5
'Printer.Print leye.ly5
'Printer.CurrentX = 0.5
'Printer.Print leye.ly6
'Printer.CurrentX = 0.5
'Printer.Print leye.ly7
'Printer.CurrentX = 0.5
'Printer.Print leye.ly8
'Printer.Print ""
'Printer.Print ""
'Printer.Line (0.5, Printer.CurrentY)-(19.7, Printer.CurrentY)
'Printer.Print ""
'Printer.Print ""
'Printer.Line (0.5, Printer.CurrentY)-(19.7, Printer.CurrentY)
'Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Line (3, Printer.CurrentY)-(7.7, Printer.CurrentY)
Printer.Line (13, Printer.CurrentY)-(18, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 0.1
Printer.CurrentX = 3.5
'Printer.Print "Firma del Directora.";
Printer.Print ini.Rector;
'Printer.Print Rector;
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, SP, profe
Close #NAR
Printer.CurrentX = 15.6 - ((Len(RTrim(profe.nombres) & " " & RTrim(profe.apellidos)) / 4.8) / 2)
Printer.Print RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
Printer.CurrentX = 4.5
Printer.Print vini.VRector;
Printer.CurrentX = 14
Printer.Print vini.VDirector

Screen.MousePointer = 0
Printer.EndDoc
End If
'FORMATO.Show 1
End Sub

Private Sub Command3_Click()
If Dir(Ruta & Combo4.Text & ".gru") = "" Then
    MsgBox "NO EXISTE ESTE GRUPO", 48, "Imprimir Reportes"
    Combo4.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo4.Text Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
        SP = RTrim(icur.director)
    End If
Wend
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, SP, profe
Close #NAR
PERI = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
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
If Combo3.Text = "PRIMERO" Then
    lwe = 1
End If
If Combo3.Text = "SEGUNDO" Then
    lwe = 2
End If
If Combo3.Text = "TERCERO" Then
    lwe = 3
End If
If Combo3.Text = "CUARTO" Then
    lwe = 4
End If
If Combo3.Text = "FINAL" Then
    lwe = 5
End If
ret = 0
Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
' *****LLAMAR FORMULARIO DE IMPRESION POR INTERVALO *****
Impr_ReportAcadem2.Frame2.Caption = Combo4.Text
Impr_ReportAcadem2.Show 1
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta por pantalla o impresión de boletines."
End Sub

Private Sub MATI20_Click()
If MATI20.Row > 0 Then
   MATI20.Col = 0
   MATI20.ToolTipText = Left(RTrim(MATI20.Text), 200)
End If
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
Private Sub Form_Load()
'Dim icur As inforcur
MATI20.Row = 0
MATI20.Col = 0
MATI20.ColWidth(0) = 8000
MATI20.CellFontBold = True
MATI20.CellForeColor = RGB(255, 255, 255)
MATI20.CellBackColor = RGB(0, 0, 150)
MATI20.Text = "                M A T E R I A"
MATI20.Col = 1
MATI20.ColWidth(1) = 600
MATI20.CellFontBold = True
MATI20.CellForeColor = RGB(255, 255, 255)
MATI20.CellBackColor = RGB(0, 0, 150)
MATI20.Text = "  %"
MATI20.Col = 2
MATI20.ColWidth(2) = 600
MATI20.CellFontBold = True
MATI20.CellForeColor = RGB(255, 255, 255)
MATI20.CellBackColor = RGB(0, 0, 150)
MATI20.Text = "DESP"

If Dir(Ruta & "conf_encabeza2.txt") <> "" Then
    Open Ruta & "conf_encabeza2.txt" For Input As #NAR
    Input #NAR, ConfTexto
    Close #NAR
End If

If Dir(Ruta & "infcur.edu") <> "" Then
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo4.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Combo4.Text = Combo4.List(0)
Else
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
End If
Text2.MaxLength = 3
End Sub
