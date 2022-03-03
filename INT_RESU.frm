VERSION 5.00
Begin VB.Form INT_RESU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervalo de impresión"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "INT_RESU.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   3855
      Begin VB.CheckBox Check6 
         Caption         =   "Si"
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Si"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Si"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ver promoción?"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fir&ma del Secretario?"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Firma del &Rector?"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "-"
         Height          =   320
         Left            =   3480
         TabIndex        =   21
         Top             =   225
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   320
         Left            =   3285
         TabIndex        =   20
         Top             =   225
         Width           =   195
      End
      Begin VB.TextBox Txt_Espa 
         Height          =   320
         Left            =   2925
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   225
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "&Encabezado"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   270
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sin &Final"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   765
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4 &Periodos"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   525
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "&Sin Formato"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&Con Formato"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      Begin VB.OptionButton Option1 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Có&digos"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicial..."
         Height          =   195
         Left            =   1320
         TabIndex        =   7
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Final..."
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   600
         Width           =   465
      End
   End
End
Attribute VB_Name = "INT_RESU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function CortaObs(Observacion As String)
Dim Recorrer As Integer, Cortar() As String, XSuma As Single
Cortar = Split(Observacion, " ")
XSuma = 0
Printer.CurrentX = 1.3
For Recorrer = 0 To UBound(Cortar)
    XSuma = XSuma + Printer.TextWidth(Cortar(Recorrer))
    If XSuma <= 17 Then
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    Else
        XSuma = Printer.TextWidth(Cortar(Recorrer))
        Printer.Print ""
        Printer.CurrentX = 1.3
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    End If
Next Recorrer
Printer.Print ""
End Function

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
    If Dir(Ruta & Frame1.Caption & argra.num_area & lwe & ".dsp") <> "" Then
        
    
        'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
        NAR = FreeFile
        VV = 0
        Open Ruta & Frame1.Caption & argra.num_area & lwe & ".dsp" For Random As #NAR Len = Len(notas_desemp)
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


Private Sub Check3_Click()
If Check3.Value = 0 Then
    Txt_Espa.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
Else
    Txt_Espa.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Dim DsFinal As String, MateriaX As Integer
If Option1.Value = True Then
    s = 1
    q = ret - 1
    MS1 = "DESEA IMPRIMIR LOS INFORMES FINALES DEL GRUPO " & Frame1.Caption & "?"
End If
If Option2.Value = True Then
    If Text1.Text = "" Then
        MsgBox "ESCRIBA EL CODIGO INICIAL", 48, "ADVERTENCIA"
        Text1.SetFocus
        Exit Sub
    End If
    If Text2.Text = "" Then
        MsgBox "ESCRIBA EL CODIGO FINAL", 48, "ADVERTENCIA"
        Text2.SetFocus
        Exit Sub
    End If
    If (Val(Text1.Text) < 1) Or (Val(Text1.Text) >= ret) Then
        MsgBox "NO EXISTE EL CODIGO INICIAL", 48, "ADVERTENCIA"
        Text1.SetFocus
        Exit Sub
    End If
    If (Val(Text2.Text) < 1) Or (Val(Text2.Text) >= ret) Then
        MsgBox "NO EXISTE EL CODIGO FINAL", 48, "ADVERTENCIA"
        Text2.SetFocus
        Exit Sub
    End If
    If Val(Text1.Text) > Val(Text2.Text) Then
        MsgBox "EL CODIGO INICIAL DEBE SER MENOR O IGUAL QUE EL FINAL", 64, "ADVERTENCIA"
        Text1.SetFocus
        Exit Sub
    End If
    s = Val(Text1.Text)
    q = Val(Text2.Text)
    MS1 = "DESEA IMPRIMIR LOS INFORMES FINALES DEL GRUPO " & Frame1.Caption & ", DESDE EL CODIGO " & Text1.Text & " HASTA EL CODIGO " & Text2.Text & "?"
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR INFORMES")
If RESP = vbYes Then
    Screen.MousePointer = 11
    If (Dir(Ruta & "rangpro.txt") <> "") And (Dir(Ruta & "promovido.txt") <> "") Then
        NAR = FreeFile
        Open Ruta & "rangpro.txt" For Input As #NAR
        Input #NAR, rus, fis
        Close #NAR
        Open Ruta & "promovido.txt" For Input As #NAR
        Input #NAR, SAPO2, SAPO3, SAPO4
        Close #NAR
    End If
    If Check2.Value = 1 Then
        If Check1.Value = 1 Then
            CX = 0.3
        Else
            CX = 0.5
        End If
    Else
        CX = 0
    End If
    Printer.ScaleMode = 7
    For VV = s To q
        'RECO = False
        Open Ruta & Frame1.Caption & ".gru" For Random As #NAR Len = Len(alugru)
        Get #NAR, VV, alugru
        Close #NAR
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        If Option4.Value = True Then
            Printer.Font.Size = 14
            Printer.CurrentY = 1
            Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
            Printer.FontBold = True
            Printer.Print ini.nombre
            Printer.CurrentX = 8.2
            Printer.Print "INFORME FINAL"
            Printer.FontBold = False
        End If
        
        If Check3.Value = 1 Then
            Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.7
        Else
            Printer.CurrentY = 2.7
        End If
        Printer.Font.Size = 10
        'If Option3.Value = True Then
        '    Printer.CurrentX = 3.7
        '    Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
        'Else
            Printer.CurrentX = 1.3
            Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
        'End If
        'If Option3.Value = True Then
        '    Printer.CurrentX = 17
        '    Printer.Print Year(Date)
        'Else
            Printer.CurrentX = 18.5
            Printer.Print "AÑO:" & Year(Date)
        'End If
        'If Option3.Value = True Then
        '    Printer.CurrentX = 2.9
        '    Printer.Print RE22;
        'Else
            Printer.CurrentX = 1.3
            Printer.Print "GRADO: " & RE22;
        'End If
        'If Option3.Value = True Then
        '    Printer.CurrentX = 7.6
        '    Printer.Print Frame1.Caption;
        'Else
            Printer.CurrentX = 6
            Printer.Print "GRUPO: " & Frame1.Caption;
        'End If
        'If Option3.Value = True Then
        '   Printer.CurrentX = 18.5
        '  Printer.Print Date
        'Else
            Printer.CurrentX = 17
            Printer.Print "FECHA: " & Date
        'End If
        Printer.Print ""
        Printer.Line (1, 4)-(20.3, 4)
        Printer.Print ""
        Printer.CurrentX = 1.3
        Printer.Print "M A T E R I A S";
        Printer.CurrentX = 7.5
        Printer.Print "I.H.";
        If Check1.Value = 1 Then
            Printer.CurrentX = 8.7
            Printer.Print "I PERIODO";
            Printer.CurrentX = 11.4 + CX
            Printer.Print "II PERIODO";
            Printer.CurrentX = 14 + (2 * CX)
            Printer.Print "III PERIODO";
            Printer.CurrentX = 16.5 + (4 * CX)
            Printer.Print "IV PERIODO";
        Else
            Printer.CurrentX = 9
            Printer.Print "I PERIODO";
            Printer.CurrentX = 12.3 + (2 * CX)
            Printer.Print "II PERIODO";
            Printer.CurrentX = 15.7 + (3 * CX)
            Printer.Print "III PERIODO";
        End If
        If Check2.Value = 0 Then
            Printer.CurrentX = 19
            Printer.Print "FINAL"
        Else
            Printer.Print ""
        End If
        Printer.Print ""
        Printer.Line (1, Printer.CurrentY)-(20.3, Printer.CurrentY)
        Printer.Print ""
            
                
        MateriaX = 0
        nf = 1
        cona = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            cona = cona + 1
            Get #NAR, cona, argra
            If RTrim(argra.nom_grup) = Frame1.Caption Then
                NAR = FreeFile
                Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                Get #NAR, argra.num_area, mate
                Close #NAR
                NAR = NAR - 1
                Printer.CurrentX = 1.3
                Printer.Print RTrim(mate.nom);
                Printer.CurrentX = 7.6
                Printer.Print argra.ih;
                ContDesemp = 0
                For ww = 1 To 5
                    DsFinal = Definitiva(ww)
                    
                    ' ***** VERIFICAMOS QUE TENGA TODAS LAS NOTAS PARA MOSTRAR LA FINAL *******
                    If Check1.Value = 1 Then
                        If DsFinal = "" And ww <> 5 Then
                            ContDesemp = 1
                        End If
                    Else
                        If DsFinal = "" And ww <> 4 And ww <> 5 Then
                            ContDesemp = 1
                        End If
                    End If
                    
                    If DsFinal <> "" And DsFinal <> "ERROR" Then
                        'MATI50.TextMatrix(nf, ww + 1) = DsFinal
                        If Check1.Value = 1 Then
                            If ww = 1 Then
                                Printer.CurrentX = 9.2
                            End If
                            If ww = 2 Then
                                Printer.CurrentX = 12 + CX
                            End If
                            If ww = 3 Then
                                Printer.CurrentX = 14.6 + (2 * CX)
                            End If
                            If ww = 4 Then
                                Printer.CurrentX = 17.2 + (4 * CX)
                            End If
                            If ww = 5 Then
                                Printer.CurrentX = 19.2
                            End If
'                            If Check2.Value = 1 And ww = 5 Then
'                                Printer.Print "";
'                            Else
'                                Printer.Print Trim(DsFinal);
'                            End If
                            If Check2.Value = 1 And ww = 5 Then
                                Printer.Print "";
                            Else
                                If ContDesemp = 1 And ww = 5 Then
                                    Printer.Print "";
                                Else
                                    Printer.Print Trim(DsFinal);
                                End If
                            End If
                        Else
                            If ww = 1 Then
                                Printer.CurrentX = 9.5
                            End If
                            If ww = 2 Then
                                Printer.CurrentX = 12.8 + (2 * CX)
                            End If
                            If ww = 3 Then
                                Printer.CurrentX = 16.4 + (3 * CX)
                            End If
                            If ww = 5 Then
                                Printer.CurrentX = 19.2
                            End If
'                            If Check2.Value = 1 And ww = 5 Then
'                                Printer.Print "";
'                            Else
'                                If ww <> 4 Then
'                                    Printer.Print Trim(DsFinal);
'                                Else
'                                    Printer.Print "";
'                                End If
'                            End If



                            If Check2.Value = 1 And ww = 5 Then
                                Printer.Print "";
                            Else
                                If ww <> 4 Then
                                
                                    If ContDesemp = 1 And ww = 5 Then
                                        Printer.Print "";
                                    Else
                                        Printer.Print Trim(DsFinal);
                                    End If
                                
                                Else
                                    Printer.Print "";
                                End If
                            End If
                            
                        End If
                        ' Se verifica el total de materias perdidas (teniendo en cuenta la nota final)
                        If (Trim(DsFinal) = Trim(confdesemp.desemp(4))) And ww = 5 Then
                            MateriaX = MateriaX + 1
                        End If
                    Else
                        Printer.Print "";
                        'MATI50.TextMatrix(nf, ww + 1) = ""
                    End If
                Next ww
                Printer.Print ""
            End If
        Wend
        Close #NAR
        
        Printer.Print ""
        Printer.Line (1, Printer.CurrentY)-(20.3, Printer.CurrentY)
        Printer.Line (1, 4)-(1, Printer.CurrentY)
        Printer.Line (7.3, 4)-(7.3, Printer.CurrentY)
        Printer.Line (8.2, 4)-(8.2, Printer.CurrentY)
        If Check1.Value = 1 Then
            If Check2.Value = 0 Then
                Printer.Line (11, 4)-(11, Printer.CurrentY)
                Printer.Line (13.7, 4)-(13.7, Printer.CurrentY)
                Printer.Line (16.3, 4)-(16.3, Printer.CurrentY)
                Printer.Line (18.8, 4)-(18.8, Printer.CurrentY)
            Else
                Printer.Line (11.2, 4)-(11.2, Printer.CurrentY)
                Printer.Line (14.2, 4)-(14.2, Printer.CurrentY)
                Printer.Line (17.2, 4)-(17.2, Printer.CurrentY)
            End If
        Else
            If Check2.Value = 0 Then
                Printer.Line (11.6, 4)-(11.6, Printer.CurrentY)
                Printer.Line (15, 4)-(15, Printer.CurrentY)
                Printer.Line (18.5, 4)-(18.5, Printer.CurrentY)
            Else
                Printer.Line (12.2, 4)-(12.2, Printer.CurrentY)
                Printer.Line (16.2, 4)-(16.2, Printer.CurrentY)
            End If
        End If
        Printer.Line (20.3, 4)-(20.3, Printer.CurrentY)
        Printer.Print ""
        ' EXPLICACION CONVENCIONES
        Printer.CurrentX = 1.3
        Printer.Font.Size = 7
        Printer.Print Trim(confdesemp.desemp(1)) & "=Desempeño Superior (";
        Printer.Print confdesemp.rango(1) + 1 & "% - 100%),  ";
        Printer.Print Trim(confdesemp.desemp(2)) & "=Desempeño Alto (";
        Printer.Print confdesemp.rango(2) + 1 & "% - " & confdesemp.rango(1) & "%),  ";
        Printer.Print Trim(confdesemp.desemp(3)) & "=Desempeño Básico (";
        Printer.Print confdesemp.rango(3) + 1 & "% - " & confdesemp.rango(2) & "%),  ";
        Printer.Print Trim(confdesemp.desemp(4));
        If Trim(confdesemp.desemp(4)) = "*LEP" Then
            Printer.Print "=Logro en Proceso (";
        Else
            Printer.Print "=Desempeño Bajo (";
        End If
        Printer.Print "0% - " & confdesemp.rango(3) & "%)."
        Printer.Font.Size = 10
        Printer.Print ""
        Printer.Print ""
        'Verificar para mostrar la información de promoción.
        If Check6.Value = 1 Then
            Printer.CurrentX = 1.3
            If MateriaX <= rus Then
                Printer.Print RTrim(SAPO2)
            End If
            If (MateriaX > rus) And (MateriaX <= fis) Then
                Printer.Print RTrim(SAPO3)
            End If
            If MateriaX > fis Then
                Printer.Print RTrim(SAPO4)
            End If
        Else
            Printer.Print ""
        End If
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = 1.3
        Printer.Print "Observaciones:"
        Printer.Print ""
        J = 0
        If (Dir(Ruta & "lrf" & Frame1.Caption & ".lrf") <> "") And (Dir(Ruta & "orf" & Frame1.Caption & ".orf") <> "") Then
            cona = 0
            Open Ruta & "lrf" & Frame1.Caption & ".lrf" For Random As #NAR Len = Len(leyfin)
            While Not EOF(NAR)
                cona = cona + 1
                Get #NAR, cona, leyfin
                If Val(leyfin.num_carnet) = Val(alugru.num_carnet) Then
                    NAR = FreeFile
                    Open Ruta & "orf" & Frame1.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
                    For I = 1 To 5
                        If leyfin.fnob(I) <> 0 Then
                            Get #NAR, leyfin.fnob(I), obsfin
                            Printer.CurrentX = 1.3
                            
                            XMax = Printer.TextWidth(Trim(obsfin))
                            If XMax > 17 Then
                                CortaObs (Trim(obsfin))
                            Else
                                Printer.Print Trim(obsfin)
                            End If
                        Else
                            J = J + 1
                        End If
                    Next I
                    Close #NAR
                    Close #(NAR - 1)
                    NAR = NAR - 1
                    GoTo saobfi
                End If
            Wend
            Close #NAR
        Else
            J = 5
        End If
saobfi:
        If J <> 0 Then
            For I = 1 To J
                Printer.Print ""
            Next I
        End If
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        'If Option2.Value = True Then
            'Y = Printer.CurrentY
            If Check4.Value = 1 Then
                'Printer.CurrentY = Y
                Printer.Line (3, Printer.CurrentY)-(8.5, Printer.CurrentY)
                Printer.CurrentY = Printer.CurrentY + 0.2
                Printer.CurrentX = 5
                Printer.Print vini.VRector;
            End If
            If Check5.Value = 1 Then
                'Printer.CurrentY = Y
                Printer.Line (13, Printer.CurrentY)-(18.5, Printer.CurrentY)
                Printer.CurrentY = Printer.CurrentY + 0.2
                Printer.CurrentX = 14.3
                Printer.Print "Firma del Secretario"
            End If
        'End If
        Printer.NewPage
    Next VV
    Printer.EndDoc
    'Printer.Font.Size = 8
    Unload Me
End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Txt_Espa.Text = Txt_Espa.Text + 1
End Sub

Private Sub Command3_Click()
Txt_Espa.Text = Txt_Espa.Text - 1
End Sub

Private Sub Form_Load()
Option1.Value = True
Option4.Value = True
Text1.MaxLength = 2
Text2.MaxLength = 2
Check4.Value = 1
Check5.Value = 1
Frame1.Caption = RESUFINA.Combo1.Text
Txt_Espa.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Option2_Click()
Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
Else
    Option2.Value = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
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
