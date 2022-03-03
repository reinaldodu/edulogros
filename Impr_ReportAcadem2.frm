VERSION 5.00
Begin VB.Form Impr_ReportAcadem2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Reportes Académicos por grupo"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4845
   Icon            =   "Impr_ReportAcadem2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   4575
      Begin VB.TextBox Text3 
         Height          =   320
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   320
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Códigos"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Todos"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
         Height          =   195
         Left            =   1920
         TabIndex        =   11
         Top             =   960
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         Height          =   320
         Left            =   3600
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   320
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Configurar margen superior"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sin Formato"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Con Formato"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Impr_ReportAcadem2"
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
     If Option2.Value = True Then
        Printer.CurrentY = 0.5
        Printer.Font.Size = 14
        Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
        Printer.FontBold = True
        Printer.Print ini.nombre
        Printer.CurrentX = 7.4
        Printer.Print "INFORME ACADÉMICO"
        Printer.FontBold = False
    End If
    
    Printer.Font.Size = 10
    Printer.CurrentY = 2.2
    'Printer.CurrentX = 5.5
    'Printer.Print Format(vini.VPeriodo, ">") & ": " & Combo3.Text
    ' ******** IMPRIME ENCABEZADO ADICIONAL DEL REPORTE DE MITAD DE TRIMESTRE *********
    Printer.CurrentX = (22 - Printer.TextWidth(ConfTexto)) / 2
    Printer.Print ConfTexto
    
    If Check1.Value = 1 Then
       Printer.CurrentY = Val(Text1.Text) / 10 + 3
    Else
       Printer.CurrentY = 3
    End If
    'Printer.CurrentY = 2.5
     
    Printer.CurrentX = 0.5
    Printer.Print Format(vini.VEstudiante, ">") & ": " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
    Printer.CurrentX = 16.5
    Printer.Print Format(vini.VFecha, ">") & ": " & Format(Format(Date, "mmm/dd/yyyy"), ">")
    Printer.CurrentX = 0.5
    Printer.Print Format(vini.VGrupo, ">") & ": " & Frame2.Caption
    Printer.CurrentX = 0.5
    'Printer.Print Format(vini.VPeriodo, ">") & ": " & Combo3.Text
    If Option2.Value = True Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Font.Size = 12
        Printer.FontBold = True
        Printer.CurrentX = 0.5
        Printer.Print "MATERIAS";
        Printer.Font.Size = 8
        Printer.CurrentX = 16
        Printer.CurrentY = Printer.CurrentY + 0.2
        Printer.Print "PORCENTAJE";
        Printer.CurrentX = 18
        Printer.Print "DESEMPEÑO"
        Printer.FontBold = False
        Printer.Line (0.5, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)
        Printer.Line (0.5, Printer.CurrentY + 0.1)-(20, Printer.CurrentY + 0.1)
    End If
    Printer.CurrentY = 5.5
End Function

Private Sub Check1_Click()
If Check1.Value = 0 Then
    Text1.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
Else
    Text1.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Text1.Text = Text1.Text + 1
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text - 1
End Sub

Private Sub Command3_Click()
Dim ValiSalto As Boolean, XMax As Single
Dim VeriPrint As Boolean
If Option3.Value = True Then
    s = 1
    q = ret - 1
    MS1 = "DESEA IMPRIMIR LOS BOLETINES DEL GRUPO " & Frame2.Caption & " DEL PERIODO " & cons_nota2.Combo3.Text & "?"
End If
If Option4.Value = True Then
    If Text2.Text = "" Then
        MsgBox "ESCRIBA EL CODIGO INICIAL", 48, "ADVERTENCIA"
        Text2.SetFocus
        Exit Sub
    End If
    If Text3.Text = "" Then
        MsgBox "ESCRIBA EL CODIGO FINAL", 48, "ADVERTENCIA"
        Text3.SetFocus
        Exit Sub
    End If
    If (Val(Text2.Text) < 1) Or (Val(Text2.Text) >= ret) Then
        MsgBox "NO EXISTE EL CODIGO INICIAL", 48, "ADVERTENCIA"
        Text2.SetFocus
        Exit Sub
        End If
    If (Val(Text3.Text) < 1) Or (Val(Text3.Text) >= ret) Then
        MsgBox "NO EXISTE EL CODIGO FINAL", 48, "ADVERTENCIA"
        Text3.SetFocus
        Exit Sub
    End If
    If Val(Text2.Text) > Val(Text3.Text) Then
        MsgBox "EL CODIGO INICIAL DEBE SER MENOR O IGUAL QUE EL FINAL", 64, "ADVERTENCIA"
        Text2.SetFocus
        Exit Sub
    End If
    s = Val(Text2.Text)
    q = Val(Text3.Text)
    MS1 = "DESEA IMPRIMIR LOS BOLETINES DEL GRUPO " & Frame2.Caption & ", DESDE EL CODIGO " & Text2.Text & " HASTA EL CODIGO " & Text3.Text & " DEL PERIODO " & cons_nota2.Combo3.Text & "?"
End If
'Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'Get #NAR, Val(alugru.num_carnet), alumno
'Close #NAR
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir Reportes")
'RESP = MsgBox("DESEA IMPRIMIR EL REPORTE DE " & Frame1.Caption & "?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR REPORTE")
If RESP = vbYes Then
    Screen.MousePointer = 11
    For VV = s To q
        VeriPrint = False
        Open Ruta & Frame2.Caption & ".gru" For Random As #NAR Len = Len(alugru)
        Get #NAR, VV, alugru
        Close #NAR
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        Call Encabezado
        cona = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
               
        If RTrim(argra.nom_grup) = Frame2.Caption Then
            OkObs = False
            OkDes = False
            
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
                        
            If Dir(Ruta & Frame2.Caption & argra.num_area & lwe & ".obs") <> "" Then
                NAR = FreeFile
                Y = 0
                Open Ruta & Frame2.Caption & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
                While Not EOF(NAR)
                    Y = Y + 1
                    Get #NAR, Y, notas
                    If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                    
                        For z = 1 To 10
                             If notas.area(z) <> 0 Then
                                 OkObs = True
                             End If
                         Next z
                        If OkObs = True Then
                        'OkObs = True
                            VeriPrint = True
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
            If Dir(Ruta & Frame2.Caption & argra.num_area & lwe & ".dsp") <> "" Then
                NAR = FreeFile
                w = 0
                Open Ruta & Frame2.Caption & argra.num_area & lwe & ".dsp" For Random As #NAR Len = Len(notas_desemp)
                While Not EOF(NAR)
                    w = w + 1
                    Get #NAR, w, notas_desemp
                    If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                    
                        For z = 1 To 10
                             If notas_desemp.porcentaje(z) <> 0 And notas_desemp.porcentaje(z) <= confdesemp.rango(3) Then
                                 OkDes = True
                             End If
                         Next z
                    
                        'OkDes = True
                        If OkObs = False And OkDes = True Then
                            VeriPrint = True
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
                
'                NAR = FreeFile
'                Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
'                For h = 1 To 14
'                    Get #NAR, h, confdesemp
'                    If Trim(argra.grado) = Trim(confdesemp.grado) Then
'                        Exit For
'                    End If
'                Next h
'                Close #NAR
'                NAR = NAR - 1
                
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
    'Printer.Print Rector;
    Printer.Print ini.Rector;
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    Get #NAR, SP, profe
    Close #NAR
    Printer.CurrentX = 15.6 - ((Len(RTrim(profe.nombres) & " " & RTrim(profe.apellidos)) / 4.8) / 2)
    Printer.Print RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
    Printer.CurrentX = 4.5
    Printer.Print vini.VRector;
    Printer.CurrentX = 14
    Printer.Print vini.VDirector
    If VeriPrint = True Then
        Printer.EndDoc
    Else
        Printer.KillDoc
    End If
Next VV
Unload Me
Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Load()
Option1.Value = True
Option3.Value = True
Text2.MaxLength = 3
Text3.MaxLength = 3
Text1.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
Else
    Option4.Value = True
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command3_Click
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
