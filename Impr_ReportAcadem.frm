VERSION 5.00
Begin VB.Form Impr_ReportAcadem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Reportes Académicos por grupo"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4845
   Icon            =   "Impr_ReportAcadem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir libro de notas"
         Height          =   195
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   320
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   320
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Códigos"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Todos"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         Height          =   195
         Left            =   3240
         TabIndex        =   6
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   450
      End
   End
End
Attribute VB_Name = "Impr_ReportAcadem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OkObs As Boolean, OkDes As Boolean
' ******* FUNCION QUE OBTIENE EL TOTAL DE LOGROS ACUMULADOS  Y EL TOTAL DE LOGROS ALCANZADOS POR PERIODOS (EJ: 5,4)*******
Private Function Total_Logros(DSPeriodo2 As Integer) As String
Dim Cont_Lgr2 As Integer, Lgr_Ttl2 As Integer, Cp2 As Integer
Lgr_Ttl = 0
Cp2 = 0
For lwe3 = 1 To DSPeriodo2
    If Dir(Ruta & fl & seri & argra.num_area & lwe3 & ".lgr") <> "" Then
    'OBTENER TOTAL DE LOGROS
        CROA2 = 0
        Cont_Lgr2 = 0
        NAR = FreeFile
        Open Ruta & fl & seri & argra.num_area & lwe3 & ".lgr" For Random As #NAR Len = Len(logru)
        While Not EOF(NAR)
            CROA2 = CROA2 + 1
            Get #NAR, CROA2, logru
            If Trim(logru.indicador) = "L" Then
                Cont_Lgr2 = Cont_Lgr2 + 1
                Lgr_Ttl2 = Lgr_Ttl2 + 1
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
        
        If Dir(Ruta & Frame2.Caption & argra.num_area & lwe3 & ".dsp") <> "" Then
            'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
            NAR = FreeFile
            VV2 = 0
            Open Ruta & Frame2.Caption & argra.num_area & lwe3 & ".dsp" For Random As #NAR Len = Len(notas_desemp)
            While Not EOF(NAR)
                VV2 = VV2 + 1
                Get #NAR, VV2, notas_desemp
                If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                    GoTo PorcentEncontrado3
                End If
            Wend
PorcentEncontrado3:
            Close #NAR
            NAR = NAR - 1
            
            For w2 = 1 To Cont_Lgr2
                'Se verifican los logros mayores al 69% (o rango menor)
                If notas_desemp.porcentaje(w2) > confdesemp.rango(3) Then
                    Cp2 = Cp2 + 1
                End If
            Next w2
        End If
        
    End If
Next lwe3
Total_Logros = Lgr_Ttl2 & "," & Cp2
End Function

' ******* FUNCION QUE OBTIENE EL DESEMPEÑO ACUMULADO POR PERIODOS Y LA DEFINITIVA (FINAL) *******
Private Function DefinitivaAcum(DsPeriodo As Integer) As String
Dim VeriManual As Boolean
Dim w As Integer, CP As Integer, PorcentLogro As Single, PromLogros As Single, SumDesemp As Long, PorcentManual(10) As Integer, ContPorcent As Integer
DefinitivaAcum = ""

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

For lwe2 = 1 To DsPeriodo
'If DsPeriodo <> lwe2 Or DsPeriodo <> 5 Then
    If Dir(Ruta & Frame2.Caption & argra.num_area & lwe2 & ".dsp") <> "" Then
        
    
        'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
        NAR = FreeFile
        VV9 = 0
        Open Ruta & Frame2.Caption & argra.num_area & lwe2 & ".dsp" For Random As #NAR Len = Len(notas_desemp)
        While Not EOF(NAR)
            VV9 = VV9 + 1
            Get #NAR, VV9, notas_desemp
            If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                GoTo PorcentEncontrado2
            End If
        Wend
PorcentEncontrado2:
        Close #NAR
        NAR = NAR - 1
        
        'OBTENER TOTAL DE LOGROS
        CROA = 0
        Cont_Lgr = 0
        NAR = FreeFile
        Open Ruta & fl & seri & argra.num_area & lwe2 & ".lgr" For Random As #NAR Len = Len(logru)
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
                'DefinitivaAcum = ""
                'Exit Function
                'Exit For
                GoTo SaltaDesemp2
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
                DefinitivaAcum = Trim(confdesemp.desemp(4))
            Else
                ' ***Si alcanza el porcentaje mínimo de logros, se promedia para obtener el desempeño
                If PromLogros <= confdesemp.rango(3) Then
                    DefinitivaAcum = Trim(confdesemp.desemp(3))
                End If
                If (PromLogros > confdesemp.rango(3)) And (PromLogros <= confdesemp.rango(2)) Then
                     DefinitivaAcum = Trim(confdesemp.desemp(3))
                End If
                
                If (PromLogros > confdesemp.rango(2)) And (PromLogros <= confdesemp.rango(1)) Then
                    DefinitivaAcum = Trim(confdesemp.desemp(2))
                End If
                If PromLogros > confdesemp.rango(1) Then
                    DefinitivaAcum = Trim(confdesemp.desemp(1))
                End If
            End If
        
        Else
            If Dir(Ruta & fl & seri & argra.num_area & lwe2 & ".ptj") = "" Then
                DefinitivaAcum = "ERROR"
                Exit Function
            Else
                NAR = FreeFile
                Open Ruta & fl & seri & argra.num_area & lwe2 & ".ptj" For Random As #NAR Len = Len(porcent_manual)
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
                'DefinitivaAcum = ""
                'Exit Function
                'Exit For
                GoTo SaltaDesemp2
            End If
            If DsPeriodo <> 5 Then
                PorcentLogro = SumDesemp / 100
                If PorcentLogro <> 0 Then
                    PorcentLogro = Format(PorcentLogro, "#")
                End If
            Else
                PorcentLogro = SumDesemp / (100 * lwe2)
                If PorcentLogro <> 0 Then
                    PorcentLogro = Format(PorcentLogro, "#")
                End If
            End If
            'REPORTE POR DESEMPEÑOS
            If PorcentLogro <= confdesemp.rango(3) Then
                DefinitivaAcum = Trim(confdesemp.desemp(4))
            End If
            If (PorcentLogro > confdesemp.rango(3)) And (PorcentLogro <= confdesemp.rango(2)) Then
                 DefinitivaAcum = Trim(confdesemp.desemp(3))
            End If
            
            If (PorcentLogro > confdesemp.rango(2)) And (PorcentLogro <= confdesemp.rango(1)) Then
                DefinitivaAcum = Trim(confdesemp.desemp(2))
            End If
            If PorcentLogro > confdesemp.rango(1) Then
                DefinitivaAcum = Trim(confdesemp.desemp(1))
            End If
        End If
    
    End If
'Else
'    Exit For
SaltaDesemp2:
'End If
Next lwe2
End Function

' *******FUNCION QUE OBTIENE EL DESEMPEÑO FINAL DE LA MATERIA*******
Private Function Definitiva(Rangos() As Byte, Desempenos() As String * 5, Porcentajes() As Byte, TtLogros As Integer, VeriLgr As Boolean) As String
Dim w As Integer, CP As Integer, PorcentLogro As Single, PromLogros As Single, SumDesemp As Long, PorcentManual(10) As Integer, ContPorcent As Integer
CP = 0
ContPorcent = 0
SumDesemp = 0
If VeriLgr = False Then
    ' ******* SI LOS PORCENTAJES SON AUTOMATICOS SE HACE LO SIGUIENTE: *******
    For w = 1 To TtLogros
        'Se verifican los logros mayores al 69% (o rango menor)
        If Porcentajes(w) > Rangos(3) Then
            CP = CP + 1
        End If
        If Porcentajes(w) = 0 Then
            ContPorcent = ContPorcent + 1
        End If
        SumDesemp = SumDesemp + Porcentajes(w)
    Next w
    'Se verifica que tenga grabado porcentajes para obtener el desempeño final
    If ContPorcent = TtLogros Then
        Definitiva = ""
        Exit Function
    End If
    PorcentLogro = (100 / TtLogros) * CP
    If PorcentLogro <> 0 Then
        PorcentLogro = Format(PorcentLogro, "#")
    End If
    PromLogros = SumDesemp / TtLogros
    If PromLogros >= 0.5 Then
        PromLogros = Format(PromLogros, "#")
    End If
    'REPORTE POR DESEMPEÑOS
    If PorcentLogro <= Rangos(3) Then
            ' ***Pierde si el porcentaje total de logros alcanzados es menor o igual al Rango inferior
        Definitiva = Trim(Desempenos(4))
    Else
        ' ***Si alcanza el porcentaje mínimo de logros, se promedia para obtener el desempeño
        If PromLogros <= Rangos(3) Then
            Definitiva = Trim(Desempenos(3))
        End If
        If (PromLogros > Rangos(3)) And (PromLogros <= Rangos(2)) Then
             Definitiva = Trim(Desempenos(3))
        End If
        
        If (PromLogros > Rangos(2)) And (PromLogros <= Rangos(1)) Then
            Definitiva = Trim(Desempenos(2))
        End If
        If PromLogros > Rangos(1) Then
            Definitiva = Trim(Desempenos(1))
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
    For w = 1 To TtLogros
        If Porcentajes(w) = 0 Then
            ContPorcent = ContPorcent + 1
        End If
        SumDesemp = SumDesemp + (Porcentajes(w) * PorcentManual(w))
    Next w
    'Se verifica que tenga grabado porcentajes para obtener el desempeño final
    If ContPorcent = TtLogros Then
        Definitiva = ""
        Exit Function
    End If
    PorcentLogro = SumDesemp / 100
    If PorcentLogro <> 0 Then
        PorcentLogro = Format(PorcentLogro, "#")
    End If
    'REPORTE POR DESEMPEÑOS
    If PorcentLogro <= Rangos(3) Then
        Definitiva = Trim(Desempenos(4))
    End If
    If (PorcentLogro > Rangos(3)) And (PorcentLogro <= Rangos(2)) Then
         Definitiva = Trim(Desempenos(3))
    End If
    
    If (PorcentLogro > Rangos(2)) And (PorcentLogro <= Rangos(1)) Then
        Definitiva = Trim(Desempenos(2))
    End If
    If PorcentLogro > Rangos(1) Then
        Definitiva = Trim(Desempenos(1))
    End If
End If

End Function

Private Function Definitiva_Imp() As String
Dim AcumulaPorcent As Byte, NotAcumula As Single
Definitiva_Imp = ""
AcumulaPorcent = 0
NotAcumula = 0
'Cont_Lgr = 0

For ww = 1 To lwe
    If Dir(Ruta & Frame2.Caption & argra.num_area & ww & ".dsp") <> "" Then
    
        'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
        NAR = FreeFile
        VV9 = 0
        Open Ruta & Frame2.Caption & argra.num_area & ww & ".dsp" For Random As #NAR Len = Len(notas_desemp)
        While Not EOF(NAR)
            VV9 = VV9 + 1
            Get #NAR, VV9, notas_desemp
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
        Open Ruta & fl & seri & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
        While Not EOF(NAR)
            CROA = CROA + 1
            Get #NAR, CROA, logru
            If Trim(logru.indicador) = "L" Then
                Cont_Lgr = Cont_Lgr + 1
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
        
        'AcumulaPorcent = 0
        'NotAcumula = 0
        For I = 1 To Cont_Lgr
            'NAR = FreeFile
            'Open Ruta & fl & seri & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
            'Get #NAR, notas_desemp.logro(I), logru
            If notas_desemp.porcentaje(I) <> 0 Then
                NAR = FreeFile
                Open Ruta & fl & seri & argra.num_area & ww & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                Get #NAR, I, porcent_manual
                Close #NAR
                NAR = NAR - 1
                AcumulaPorcent = AcumulaPorcent + porcent_manual.porcent_logro
                NotAcumula = NotAcumula + (Val(porcent_manual.porcent_logro) * Val(notas_desemp.porcentaje(I)))
            End If
            'Close #NAR
            'NAR = NAR - 1
        Next I
        Definitiva_Imp = AcumulaPorcent & "," & NotAcumula
    End If
Next ww
End Function

Private Function CortaObs(Observacion As String)
Dim Recorrer As Integer, Cortar() As String, XSuma As Single
Cortar = Split(Observacion, " ")
XSuma = 0
Printer.CurrentX = 0.5
For Recorrer = 0 To UBound(Cortar)
    XSuma = XSuma + Printer.TextWidth(Cortar(Recorrer))
    If XSuma <= 16 Then
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
'If (Printer.CurrentY > 26) Then
If (Printer.CurrentY > 31) Then
    Printer.Line (18.9, 4.5)-(18.9, Printer.CurrentY)
    Printer.Line (19.9, 4.5)-(19.9, Printer.CurrentY)
    Printer.Line (21, 4.5)-(21, Printer.CurrentY)
   Printer.NewPage
   Call Encabezado
End If
End Function

Private Function Encabezado()
'Printer.ScaleMode = 7
''Hoja oficio
'Printer.PaperSize = 5
'If Printer.Page = 1 Then
'    'Imprimir logotipos
'    If CONS_NOTA.Image1.Picture <> 0 Then
'        Printer.PaintPicture CONS_NOTA.Image1.Picture, 0.5, 1, 2, 2
'    End If
'    If CONS_NOTA.Image2.Picture <> 0 Then
'        Printer.PaintPicture CONS_NOTA.Image2.Picture, 19, 1, 2, 2
'    End If
'    'Imprimir encabezado
'    Printer.CurrentY = 1
'    Printer.FontName = "Monotype corsiva"
'    Printer.FontSize = 16
'    Printer.FontBold = True
'    Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.nombre))) / 2
'    Printer.Print ini.nombre
'
'    Printer.FontSize = 10
'    Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.ciudad))) / 2
'    Printer.Print ini.ciudad
'
'    Printer.FontName = ""
'    Printer.Print ""
'
'    Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.modalidad) & " PERIODO " & CONS_NOTA.Combo3.Text & " - " & Trim(ini.Telefono))) / 2
'    Printer.Print Trim(ini.modalidad) & " PERIODO " & CONS_NOTA.Combo3.Text & " - " & Trim(ini.Telefono)
'    Printer.Print ""
'
'    'Imprimir el nombre del estudiante
'    Printer.CurrentX = (21 - Printer.TextWidth(RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").")) / 2
'    Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ")."
'
'    'Imprimir el nombre del grupo
'    Printer.CurrentX = (21 - Printer.TextWidth(Frame2.Caption)) / 2
'    Printer.Print Right(Frame2.Caption, Len(Frame2.Caption) - 1)
'
'    Printer.FontBold = False
'Else
'    Printer.CurrentY = 2.5
'    Printer.FontSize = 10
'    Printer.FontBold = True
'
'    Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.modalidad) & " PERIODO " & CONS_NOTA.Combo3.Text & " - " & Trim(ini.Telefono))) / 2
'    Printer.Print Trim(ini.modalidad) & " PERIODO " & CONS_NOTA.Combo3.Text & " - " & Trim(ini.Telefono)
'    Printer.Print ""
'
'    'Imprimir el nombre del estudiante
'    Printer.CurrentX = (21 - Printer.TextWidth(RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").")) / 2
'    Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ")."
'
'    'Imprimir el nombre del grupo
'    Printer.CurrentX = (21 - Printer.TextWidth(Frame2.Caption)) / 2
'    Printer.Print Right(Frame2.Caption, Len(Frame2.Caption) - 1)
'
'    Printer.FontBold = False
'End If
'
'Printer.Print ""
'Printer.CurrentX = 19
'Printer.Print "  %";
'Printer.CurrentX = 20
'Printer.Print " CAL"
'Printer.Line (0.5, Printer.CurrentY)-(21, Printer.CurrentY)

'********NUEVO****************
Printer.ScaleMode = 7
'Hoja oficio
Printer.PaperSize = 5
If Printer.Page = 1 Then
    'Imprimir logotipos
    If CONS_NOTA.Image1.Picture <> 0 Then
        Printer.PaintPicture CONS_NOTA.Image1.Picture, 0.5, 1, 2, 2
    End If
    If CONS_NOTA.Image2.Picture <> 0 Then
        Printer.PaintPicture CONS_NOTA.Image2.Picture, 19, 1, 2, 2
    End If
    'Imprimir encabezado
    Printer.CurrentY = 1
    Printer.FontName = "Monotype corsiva"
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.nombre))) / 2
    Printer.Print ini.nombre
    
    Printer.FontSize = 10
    Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.ciudad))) / 2
    Printer.Print ini.ciudad
    
    Printer.FontName = ""
    Printer.Print ""
    If (CONS_NOTA.Combo3.Text <> "FINAL") Then
        Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.modalidad) & " PERIODO " & CONS_NOTA.Combo3.Text & " - " & Trim(ini.Telefono))) / 2
        Printer.Print Trim(ini.modalidad) & " PERIODO " & CONS_NOTA.Combo3.Text & " - " & Trim(ini.Telefono)
        Printer.Print ""
    Else
        Printer.CurrentX = (21 - Printer.TextWidth("PROMOCIÓN FINAL")) / 2
        Printer.Print "PROMOCIÓN FINAL"
        Printer.Print ""
        '****Si se imprimen los libros de notas sale Folio No.*****
        If Check1.Value = 1 Then
            Printer.CurrentX = (21 - Printer.TextWidth("AÑO LECTIVO " & Trim(ini.Telefono))) / 2
            Printer.Print "AÑO LECTIVO " & Trim(ini.Telefono);
            Printer.CurrentX = 17.5
            Printer.Print "FOLIO No."
        Else
            Printer.CurrentX = (21 - Printer.TextWidth("AÑO LECTIVO " & Trim(ini.Telefono))) / 2
            Printer.Print "AÑO LECTIVO " & Trim(ini.Telefono)
        End If
        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
        Printer.Line (1, Printer.CurrentY + 0.07)-(21, Printer.CurrentY + 0.07)
        Printer.CurrentY = Printer.CurrentY + 0.07
        'Printer.Print ""
    End If
    If (CONS_NOTA.Combo3.Text <> "FINAL") Then
        'Imprimir el nombre del estudiante
        Printer.CurrentX = (21 - Printer.TextWidth(RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").")) / 2
        Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ")."
        
        'Imprimir el nombre del grupo
        Printer.CurrentX = (21 - Printer.TextWidth(Frame2.Caption)) / 2
        Printer.Print Right(Frame2.Caption, Len(Frame2.Caption) - 1)
    Else
        'Imprimir el nombre del estudiante
        Printer.FontSize = 12
        Printer.CurrentX = (21 - Printer.TextWidth(RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").")) / 2
        Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ")."
        
        'Imprimir el nombre del grupo
        Printer.FontSize = 10
        Printer.CurrentX = (21 - Printer.TextWidth(Frame2.Caption)) / 2
        Printer.Print Right(Frame2.Caption, Len(Frame2.Caption) - 1)
        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
        Printer.Line (1, Printer.CurrentY + 0.07)-(21, Printer.CurrentY + 0.07)
    End If
    Printer.FontBold = False
Else
    Printer.CurrentY = 2.5
    Printer.FontSize = 10
    Printer.FontBold = True
    If (CONS_NOTA.Combo3.Text <> "FINAL") Then
        Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.modalidad) & " PERIODO " & CONS_NOTA.Combo3.Text & " - " & Trim(ini.Telefono))) / 2
        Printer.Print Trim(ini.modalidad) & " PERIODO " & CONS_NOTA.Combo3.Text & " - " & Trim(ini.Telefono)
        Printer.Print ""
    Else
        Printer.CurrentX = (21 - Printer.TextWidth("PROMOCIÓN FINAL")) / 2
        Printer.Print "PROMOCIÓN FINAL"
        Printer.Print ""
        '****Si se imprimen los libros de notas sale Folio No.*****
        If Check1.Value = 1 Then
            Printer.CurrentX = (21 - Printer.TextWidth("AÑO LECTIVO " & Trim(ini.Telefono))) / 2
            Printer.Print "AÑO LECTIVO " & Trim(ini.Telefono);
            Printer.CurrentX = 17.5
            Printer.Print "FOLIO No."
        Else
            Printer.CurrentX = (21 - Printer.TextWidth("AÑO LECTIVO " & Trim(ini.Telefono))) / 2
            Printer.Print "AÑO LECTIVO " & Trim(ini.Telefono)
        End If
        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
        Printer.Line (1, Printer.CurrentY + 0.07)-(21, Printer.CurrentY + 0.07)
        Printer.CurrentY = Printer.CurrentY + 0.07
        'Printer.Print ""
    End If
    If (CONS_NOTA.Combo3.Text <> "FINAL") Then
        'Imprimir el nombre del estudiante
        Printer.CurrentX = (21 - Printer.TextWidth(RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").")) / 2
        Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ")."
        
        'Imprimir el nombre del grupo
        Printer.CurrentX = (21 - Printer.TextWidth(Frame2.Caption)) / 2
        Printer.Print Right(Frame2.Caption, Len(Frame2.Caption) - 1)
    Else
        'Imprimir el nombre del estudiante
        Printer.FontSize = 12
        Printer.CurrentX = (21 - Printer.TextWidth(RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").")) / 2
        Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ")."
        
        'Imprimir el nombre del grupo
        Printer.FontSize = 10
        Printer.CurrentX = (21 - Printer.TextWidth(Frame2.Caption)) / 2
        Printer.Print Right(Frame2.Caption, Len(Frame2.Caption) - 1)
        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
        Printer.Line (1, Printer.CurrentY + 0.07)-(21, Printer.CurrentY + 0.07)
    End If
    
    Printer.FontBold = False
End If
Printer.Print ""
If (CONS_NOTA.Combo3.Text <> "FINAL") Then
    Printer.FontBold = False
    Printer.CurrentX = 19
    Printer.Print "  %";
    Printer.CurrentX = 20
    Printer.Print " CAL"
Else
    Printer.FontSize = 9
    Printer.FontBold = True
    Printer.CurrentX = 1
    Printer.Print "ASIGNATURAS / CONCEPTOS";
    Printer.CurrentX = 16
    Printer.Print "I.H.";
    Printer.CurrentX = 17
    Printer.Print "VAL";
    Printer.CurrentX = 18.5
    Printer.Print "DESEMPEÑO"
End If
If (CONS_NOTA.Combo3.Text <> "FINAL") Then
    Printer.Line (0.5, Printer.CurrentY)-(21, Printer.CurrentY)
Else
    Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
End If

End Function

Private Sub Command3_Click()
Dim ValiSalto As Boolean, XMax As Single, GuardaY As Single
'Dim AcumulaY As Single, AcumulaPorcent As Byte, NotAcumula As Single, Actual_Y As Single
Dim Cont_Sup As Byte, Cont_Alt As Byte, Cont_Bas As Byte, Cont_Baj As Byte

If Option3.Value = True Then
    s = 1
    q = ret - 1
    If (CONS_NOTA.Combo3.Text <> "FINAL") Then
        MS1 = "DESEA IMPRIMIR LOS BOLETINES DEL GRUPO " & Frame2.Caption & " DEL PERIODO " & CONS_NOTA.Combo3.Text & "?"
    Else
        MS1 = "DESEA IMPRIMIR LOS INFORMES FINALES DEL GRUPO " & Frame2.Caption & "?"
    End If
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
    If (CONS_NOTA.Combo3.Text <> "FINAL") Then
        MS1 = "DESEA IMPRIMIR LOS BOLETINES DEL GRUPO " & Frame2.Caption & ", DESDE EL CODIGO " & Text2.Text & " HASTA EL CODIGO " & Text3.Text & " DEL PERIODO " & CONS_NOTA.Combo3.Text & "?"
    Else
        MS1 = "DESEA IMPRIMIR LOS INFORMES FINALES DEL GRUPO " & Frame2.Caption & ", DESDE EL CODIGO " & Text2.Text & " HASTA EL CODIGO " & Text3.Text & "?"
    End If
End If
'Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'Get #NAR, Val(alugru.num_carnet), alumno
'Close #NAR
If (CONS_NOTA.Combo3.Text <> "FINAL") Then
    RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir Reportes")
    'RESP = MsgBox("DESEA IMPRIMIR EL REPORTE DE " & Frame1.Caption & "?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR REPORTE")
    If RESP = vbYes Then
        Screen.MousePointer = 11
        For VV = s To q
            Cont_Sup = 0
            Cont_Alt = 0
            Cont_Bas = 0
            Cont_Baj = 0
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
                    If Dir(Ruta & Frame2.Caption & argra.num_area & lwe & ".obs") <> "" Then
                        NAR = FreeFile
                        Y = 0
                        Open Ruta & Frame2.Caption & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
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
                                Call SaltaLinea
                                Printer.CurrentX = 0.5
                                Printer.FontSize = 9
                                Printer.FontBold = True
                                Impr_adicional = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:" & notas.FA & ")"
                                Printer.Print Impr_adicional;
                                Printer.FontSize = 8
                                Printer.CurrentX = 14.2
                                If Trim(mate.nom) <> "CONVIVENCIA ESCOLAR" Then
                                    Printer.Print "Valoración promedio acumulada:";
                                    'Call Definitiva_Imp
                                    ValPromAcu = Definitiva_Imp
                                    If ValPromAcu <> "" Then
                                        ValPromAcu = Split(ValPromAcu, ",")
                                        If ValPromAcu(0) <> 0 Then
                                            Printer.CurrentX = 19.2
                                            Printer.Print ValPromAcu(0);
                                            Printer.CurrentX = 20
                                            Printer.Print Format(ValPromAcu(1) / ValPromAcu(0), "#.00")
                                        Else
                                            Printer.Print "";
                                            Printer.Print ""
                                        End If
                                    Else
                                        Printer.Print ""
                                    End If
                                Else
                                    Printer.Print ""
                                End If
                                Printer.FontBold = False
                                GoTo encontrar
                            End If
                        Wend
encontrar:
                        Close #NAR
                        NAR = NAR - 1
                    Else
                        notas.FA = 0
                    End If
                    If Dir(Ruta & Frame2.Caption & argra.num_area & lwe & ".dsp") <> "" Then
                        NAR = FreeFile
                        VV7 = 0
                        Open Ruta & Frame2.Caption & argra.num_area & lwe & ".dsp" For Random As #NAR Len = Len(notas_desemp)
                        While Not EOF(NAR)
                            VV7 = VV7 + 1
                            Get #NAR, VV7, notas_desemp
                            If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                                OkDes = True
                                If OkObs = False Then
                                    NAR = FreeFile
                                    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                                    Get #NAR, argra.num_area, mate
                                    Close #NAR
                                    NAR = NAR - 1
                                    Call SaltaLinea
                                    Printer.CurrentX = 0.5
                                    Printer.FontSize = 9
                                    Printer.FontBold = True
                                    Impr_adicional = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:" & notas.FA & ")"
                                    Printer.Print Impr_adicional;
                                    
                                    Printer.FontSize = 8
                                    Printer.CurrentX = 14.2
                                    If Trim(mate.nom) <> "CONVIVENCIA ESCOLAR" Then
                                        Printer.Print "Valoración promedio acumulada:";
                                        'Call Definitiva_Imp
                                        ValPromAcu = Definitiva_Imp
                                        If ValPromAcu <> "" Then
                                            ValPromAcu = Split(ValPromAcu, ",")
                                            If ValPromAcu(0) <> 0 Then
                                                Printer.CurrentX = 19.2
                                                Printer.Print ValPromAcu(0);
                                                Printer.CurrentX = 20
                                                Printer.Print Format(ValPromAcu(1) / ValPromAcu(0), "#.00")
                                            Else
                                                Printer.Print "";
                                                Printer.Print ""
                                            End If
                                        Else
                                            Printer.Print ""
                                        End If
                                    Else
                                        Printer.Print ""
                                    End If
                                    Printer.FontBold = False
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
                        'AcumulaPorcent = 0
                        'NotAcumula = 0
                        For I = 1 To Cont_Lgr
                            NAR = FreeFile
                            Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
                            For h = 1 To 14
                                Get #NAR, h, confdesemp
                                If Trim(argra.grado) = Trim(confdesemp.grado) Then
                                    Exit For
                                End If
                            Next h
                            Close #NAR
            '                'VALIDA RANGOS POR DESEMPEÑO
                             If notas_desemp.porcentaje(I) <> 0 Then
                                'Call SaltaLinea
                                'Printer.FontSize = 8
                                'Printer.CurrentX = 18.8
                                If notas_desemp.porcentaje(I) <= confdesemp.rango(3) Then
                                    Cont_Baj = Cont_Baj + 1
            '                        If notas_desemp.recuperado(I) = False Then
            '                            Printer.Print confdesemp.desemp(4);
            '                        Else
            '                            Printer.Print confdesemp.recupera(4);
            '                        End If
                                End If
                                If (notas_desemp.porcentaje(I) > confdesemp.rango(3)) And (notas_desemp.porcentaje(I) <= confdesemp.rango(2)) Then
                                    Cont_Bas = Cont_Bas + 1
            '                        If notas_desemp.recuperado(I) = False Then
            '                            Printer.Print confdesemp.desemp(3);
            '                        Else
            '                            Printer.Print confdesemp.recupera(3);
            '                        End If
                                End If
            
                                If (notas_desemp.porcentaje(I) > confdesemp.rango(2)) And (notas_desemp.porcentaje(I) <= confdesemp.rango(1)) Then
                                    Cont_Alt = Cont_Alt + 1
            '                       If notas_desemp.recuperado(I) = False Then
            '                            Printer.Print confdesemp.desemp(2);
            '                        Else
            '                            Printer.Print confdesemp.recupera(2);
            '                        End If
                                End If
                                If notas_desemp.porcentaje(I) > confdesemp.rango(1) Then
                                    Cont_Sup = Cont_Sup + 1
            '                        If notas_desemp.recuperado(I) = False Then
            '                            Printer.Print confdesemp.desemp(1);
            '                        Else
            '                            Printer.Print confdesemp.recupera(1);
            '                        End If
                                End If
                            End If
                            NAR = FreeFile
                            Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                            Get #NAR, notas_desemp.logro(I), logru
                            If notas_desemp.porcentaje(I) <> 0 Then
                                NAR = FreeFile
                                Open Ruta & fl & seri & argra.num_area & lwe & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                                Get #NAR, I, porcent_manual
                                Close #NAR
                                NAR = NAR - 1
                                Call SaltaLinea
                                Printer.FontSize = 8
                                Printer.CurrentX = 19.2
                                Printer.Print porcent_manual.porcent_logro;
                                'AcumulaPorcent = AcumulaPorcent + porcent_manual.porcent_logro
                                Printer.CurrentX = 20.2
                                Printer.Print notas_desemp.porcentaje(I);
                                'NotAcumula = NotAcumula + (Val(porcent_manual.porcent_logro) * Val(notas_desemp.porcentaje(I)))
                                '*** SE VERIFICA EL TAMAÑO DEL LOGRO PARA EL SALTO DE LINEA ***
                                XMax = Printer.TextWidth(Trim(logru.indicador) & " - " & Trim(logru.observ))
                                If XMax > 16 Then
                                    CortaObs (Trim(logru.indicador) & " - " & Trim(logru.observ))
                                Else
                                    'Call SaltaLinea
                                    Printer.CurrentX = 0.5
                                    Printer.FontSize = 8
                                    Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                                End If
                            End If
                            Close #NAR
                            NAR = NAR - 1
                        Next I
                    End If
                    '****IMPRIMIR OBSERVACIONES ******
                    If OkObs = True Then
                        For I = 1 To 10
                            If notas.area(I) <> 0 Then
                                NAR = FreeFile
                                Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                                Get #NAR, notas.area(I), logru
                                'If Trim(logru.indicador) <> "L" And Trim(logru.indicador) <> "O" And Trim(logru.indicador) <> "S" Then
                                If Trim(logru.indicador) <> "L" Then
                                    '*** SE VERIFICA EL TAMAÑO DEL LOGRO PARA EL SALTO DE LINEA ***
                                    XMax = Printer.TextWidth(Trim(logru.indicador) & " - " & Trim(logru.observ))
                                    If XMax > 16 Then
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
                    End If
                    If OkDes = True Or OkObs = True Then
                        Printer.Print ""
                    End If
                Printer.Line (0.5, Printer.CurrentY)-(21, Printer.CurrentY)
                End If
            Wend
            Close #NAR
            ' ********* IMPRIMIR PIE *********
            Printer.Line (18.9, 4.5)-(18.9, Printer.CurrentY)
            Printer.Line (19.9, 4.5)-(19.9, Printer.CurrentY)
            Printer.Line (21, 4.5)-(21, Printer.CurrentY)
            
            ' VERIFICAR TAMAÑO NECESARIO PARA IMPRIMIR PIE
            'If Printer.CurrentY > 19.5 Then
            If Printer.CurrentY > 24.5 Then
               Printer.NewPage
               Call Encabezado
            End If
            Printer.Print ""
            'Printer.Print ""
            Open Ruta & "leyenda.edu" For Input As #NAR
            Input #NAR, leye.ly1, leye.ly2, leye.ly3, leye.ly4, leye.ly5, leye.ly6, leye.ly7, leye.ly8
            Close #NAR
            Printer.Font.Size = 9
            If Trim(leye.ly1) <> "" Then
                Printer.CurrentX = 0.5
                Printer.Print leye.ly1
            End If
            If Trim(leye.ly2) <> "" Then
                Printer.CurrentX = 0.5
                Printer.Print leye.ly2
            End If
            If Trim(leye.ly3) <> "" Then
                Printer.CurrentX = 0.5
                Printer.Print leye.ly3
            End If
            If Trim(leye.ly4) <> "" Then
                Printer.CurrentX = 0.5
                Printer.Print leye.ly4
            End If
            If Trim(leye.ly5) <> "" Then
                Printer.CurrentX = 0.5
                Printer.Print leye.ly5
            End If
            If Trim(leye.ly6) <> "" Then
                Printer.CurrentX = 0.5
                Printer.Print leye.ly6
            End If
            Printer.Print ""
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
            Printer.Print "TOTAL LOGROS"
            Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
            Printer.FontBold = False
            Printer.CurrentX = 0.7
            Printer.Print confdesemp.desemp(1) & "(Desempeño Superior)";
            Printer.CurrentX = 5.5
            Printer.Print confdesemp.rango(1) + 1 & "% - 100%";
            Printer.CurrentX = 8.8
            'Printer.Print confdesemp.recupera(1)
            'Printer.Print ""
            Printer.Print Cont_Sup
            Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
            Printer.CurrentX = 0.7
            Printer.Print confdesemp.desemp(2) & "(Desempeño Alto)";
            Printer.CurrentX = 5.5
            Printer.Print confdesemp.rango(2) + 1 & "% - " & confdesemp.rango(1) & "%";
            Printer.CurrentX = 8.8
            'Printer.Print confdesemp.recupera(2)
            'Printer.Print ""
            Printer.Print Cont_Alt
            Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
            Printer.CurrentX = 0.7
            Printer.Print confdesemp.desemp(3) & "(Desempeño Básico)";
            Printer.CurrentX = 5.5
            Printer.Print confdesemp.rango(3) + 1 & "% - " & confdesemp.rango(2) & "%";
            Printer.CurrentX = 8.8
            'Printer.Print confdesemp.recupera(3)
            'Printer.Print ""
            Printer.Print Cont_Bas
            Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
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
            'Printer.Print confdesemp.recupera(4)
            'Printer.Print ""
            Printer.Print Cont_Baj
            Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
            'Trazar líneas verticales
            Printer.Line (0.5, GuardaY)-(0.5, Printer.CurrentY)
            Printer.Line (5, GuardaY)-(5, Printer.CurrentY)
            Printer.Line (7.8, GuardaY)-(7.8, Printer.CurrentY)
            Printer.Line (10.6, GuardaY)-(10.6, Printer.CurrentY)
            Printer.FontSize = 9
            If Trim(leye.ly7) <> "" Then
                Printer.CurrentX = 0.5
                Printer.Print leye.ly7
            End If
            If Trim(leye.ly8) <> "" Then
                Printer.CurrentX = 0.5
                Printer.Print leye.ly8
            End If
            Printer.Print ""
            Printer.Print ""
            Printer.Print ""
            Printer.Print ""
            Printer.Line (3, Printer.CurrentY)-(7.7, Printer.CurrentY)
            Printer.Line (13, Printer.CurrentY)-(18, Printer.CurrentY)
            Printer.CurrentY = Printer.CurrentY + 0.1
            Printer.CurrentX = 3.5
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
            Screen.MousePointer = 0
            'Printer.NewPage
            Printer.EndDoc
    Next VV
    'Printer.EndDoc
    Printer.PaperSize = 1
    Unload Me
    Screen.MousePointer = 0
    End If
    
Else
'Else

' **********  IMPRESIÓN DEL REPORTE FINAL  ****************
'**********************************************************
    lwe = 4
    cona = 0
    RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir informes finales")
    If RESP = vbYes Then
        Screen.MousePointer = 11
        NAR = FreeFile
        For VV = s To q
            'NAR = FreeFile
            Open Ruta & Frame2.Caption & ".gru" For Random As #NAR Len = Len(alugru)
            Get #NAR, VV, alugru
            Close #NAR
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, (Val(alugru.num_carnet)), alumno
            Close #NAR
            Call Encabezado
            MateriaX = 0
            VeriNivela = False
            NivelaTXT = ""
            'NAR = FreeFile
            cona = 0
            Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
            While Not EOF(NAR)
                cona = cona + 1
                Get #NAR, cona, argra
                If RTrim(argra.nom_grup) = Frame2.Caption Then
                    NAR = FreeFile
                    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                    Get #NAR, argra.num_area, mate
                    Close #NAR
                    Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
                    For h = 1 To 14
                        Get #NAR, h, confdesemp
                        If Trim(argra.grado) = Trim(confdesemp.grado) Then
                            Exit For
                        End If
                    Next h
                    Close #NAR
                    NAR = NAR - 1
                    'MATI20.Rows = MATI20.Rows + 1
                    'MATI20.Col = 0
                    'MATI20.Row = MATI20.Rows - 1
                    'MATI20.CellFontBold = True
                    'MATI20.CellForeColor = RGB(0, 0, 255)
                    'MATI20.TextMatrix(MATI20.Rows - 1, 0) = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:0)"
                    
                    
                    '****CONVIVENCIA ESCOLAR(27) SE DISCRIMINA EN DISCIPLINA Y CONDUCTA******
                    If (mate.num <> 27) Then
                        Printer.CurrentX = 1
                        Printer.FontSize = 9
                        Printer.FontBold = True
                        Printer.Print RTrim(mate.nom);
                        Printer.FontBold = False
                        Printer.CurrentX = 16
                        Printer.Print argra.ih;
                        Printer.CurrentX = 17
                        'MATI20.TextMatrix(MATI20.Rows - 1, 0) = RTrim(mate.nom) & "  (I.H:" & argra.ih & ")"
                        'RowDesemp = MATI20.Rows - 1
                        ValPromAcu = Definitiva_Imp
                        If ValPromAcu <> "" Then
                            ValPromAcu = Split(ValPromAcu, ",")
                            If ValPromAcu(0) <> 0 Then
                                'MATI20.Col = 1
                                'MATI20.Row = MATI20.Rows - 1
                                'MATI20.CellFontBold = True
                                'MATI20.TextMatrix(MATI20.Rows - 1, 1) = ValPromAcu(0)
                                Printer.Print Format(ValPromAcu(1) / ValPromAcu(0), "#.00");
                                'MATI20.Col = 2
                                'VALIDA RANGOS POR DESEMPEÑO
                                Printer.CurrentX = 18.5
                                If (ValPromAcu(1) / ValPromAcu(0)) <> "" Then
                                    If (ValPromAcu(1) / ValPromAcu(0)) <= confdesemp.rango(3) Then
                                        'MATI20.CellFontBold = True
                                        MateriaX = MateriaX + 1
                                        
                                        '**********SE VERIFICA SI LA MATERIA FUE NIVELADA*********
                                        If Dir(Ruta & Frame2.Caption & argra.num_area & "5.dsp") <> "" Then
                                            'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
                                            NAR = FreeFile
                                            VV4 = 0
                                            Open Ruta & Frame2.Caption & argra.num_area & "5.dsp" For Random As #NAR Len = Len(notas_desemp)
                                            While Not EOF(NAR)
                                                VV4 = VV4 + 1
                                                Get #NAR, VV4, notas_desemp
                                                If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                                                    If (notas_desemp.porcentaje(1) <> 0) Or (notas_desemp.porcentaje(2) <> 0) Then
                                                        VeriNivela = True
                                                        For I = 1 To 2
                                                            If (notas_desemp.porcentaje(I) <> 0) Then
                                                                NivelaTXT = NivelaTXT + RTrim(mate.nom) & "$" & notas_desemp.porcentaje(I)
                                                                'OBTENER CODIGO DE OBSERVACIÓN
                                                                If Dir(Ruta & Frame2.Caption & argra.num_area & "5.obs") <> "" Then
                                                                    NAR = FreeFile
                                                                    VV2 = 0
                                                                    Open Ruta & Frame2.Caption & argra.num_area & "5.obs" For Random As #NAR Len = Len(notas)
                                                                    While Not EOF(NAR)
                                                                        VV2 = VV2 + 1
                                                                        Get #NAR, VV2, notas
                                                                        If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                                                                            'For I = 1 To 2
                                                                                If notas.area(I) <> 0 Then
                                                                                    NAR = FreeFile
                                                                                    Open Ruta & fl & seri & argra.num_area & "5.lgr" For Random As #NAR Len = Len(logru)
                                                                                    Get #NAR, notas.area(I), logru
                                                                                        'If (notas_desemp.porcentaje(I) <> 0) Then
                                                                                            NivelaTXT = NivelaTXT + "$" & RTrim(logru.observ)
                                                                                        'End If
                                                                                    Close #NAR
                                                                                    NAR = NAR - 1
                                                                                Else
                                                                                    'NivelaTXT = NivelaTXT + "%"
                                                                                End If
                                                                            'Next I
                                                                            GoTo PorcentEncontrado22
                                                                        'Else
                                                                            'NivelaTXT = NivelaTXT + "%"
                                                                        End If
                                                                    Wend
PorcentEncontrado22:
                                                                    
                                                                    Close #NAR
                                                                    NAR = NAR - 1
                                                                Else
                                                                    'NivelaTXT = NivelaTXT + "%"
                                                                End If
                                                                NivelaTXT = NivelaTXT + "%"
                                                            End If
                                                        Next I
                                                        If (notas_desemp.porcentaje(1) >= 70) Then
                                                            MateriaX = MateriaX - 1
                                                        Else
                                                            If (notas_desemp.porcentaje(1) <> 0) Then
                                                                MateriaX = MateriaX + 1
                                                            End If
                                                        End If
                                                        If (notas_desemp.porcentaje(2) >= 70) Then
                                                            MateriaX = MateriaX - 2
                                                        Else
                                                            If (notas_desemp.porcentaje(2) <> 0) Then
                                                                MateriaX = MateriaX + 2
                                                            End If
                                                        End If
                                                    End If
                                                    GoTo PorcentEncontrado11
                                                End If
                                            Wend
PorcentEncontrado11:
                                            Close #NAR
                                            NAR = NAR - 1
                                            
                                            
                                        End If
                                        
                                        Printer.Print "BAJO"
                                        Printer.CurrentX = 1
                                        Printer.FontSize = 8
                                        Printer.Print comdpe.bajo
                                        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                    End If
                                    If ((ValPromAcu(1) / ValPromAcu(0)) > confdesemp.rango(3)) And ((ValPromAcu(1) / ValPromAcu(0)) <= confdesemp.rango(2)) Then
                                        Printer.Print "BÁSICO"
                                        Printer.CurrentX = 1
                                        Printer.FontSize = 8
                                        Printer.Print comdpe.basico
                                        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                    End If
                            
                                    If ((ValPromAcu(1) / ValPromAcu(0)) > confdesemp.rango(2)) And ((ValPromAcu(1) / ValPromAcu(0)) <= confdesemp.rango(1)) Then
                                        Printer.Print "ALTO"
                                        Printer.CurrentX = 1
                                        Printer.FontSize = 8
                                        Printer.Print comdpe.alto
                                        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                    End If
                                    If (ValPromAcu(1) / ValPromAcu(0)) > confdesemp.rango(1) Then
                                        Printer.Print "SUPERIOR"
                                        Printer.CurrentX = 1
                                        Printer.FontSize = 8
                                        Printer.Print comdpe.superior
                                        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                    End If
                                End If
                            Else
                                Printer.Print ""
                                'MATI20.TextMatrix(MATI20.Rows - 1, 2) = ""
                            End If
                        End If
                    Else
                        '****** SE IMPRIME DISCIPLINA Y CONDUCTA******
                        If Dir(Ruta & Frame2.Caption & argra.num_area & "4.dsp") <> "" Then
                            NAR = FreeFile
                            VV5 = 0
                            Open Ruta & Frame2.Caption & argra.num_area & "4.dsp" For Random As #NAR Len = Len(notas_desemp)
                            While Not EOF(NAR)
                                VV5 = VV5 + 1
                                Get #NAR, VV5, notas_desemp
                                If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                                    If (Dir(Ruta & fl & seri & argra.num_area & "4.lgr") <> "") Then
                                        NAR = FreeFile
                                        Open Ruta & fl & seri & argra.num_area & "4.lgr" For Random As #NAR Len = Len(logru)
                                        For TT = 1 To 2
                                            Get #NAR, notas_desemp.logro(TT), logru
                                            If (notas_desemp.porcentaje(TT) <> 0) Then
                                                Printer.CurrentX = 1
                                                Printer.FontSize = 9
                                                Printer.FontBold = True
                                                Printer.Print Format(RTrim(logru.observ), ">");
                                                Printer.FontBold = False
                                                Printer.CurrentX = 17
                                                Printer.Print Format(notas_desemp.porcentaje(TT), "#.00");
                                                Printer.CurrentX = 18.5
                                                'RANGO DE DESEMPEÑOS
                                                If notas_desemp.porcentaje(TT) <= confdesemp.rango(3) Then
                                                    Printer.Print "BAJO"
                                                End If
                                                If (notas_desemp.porcentaje(TT) > confdesemp.rango(3)) And (notas_desemp.porcentaje(TT) <= confdesemp.rango(2)) Then
                                                    Printer.Print "BÁSICO"
                                                End If
                                        
                                                If (notas_desemp.porcentaje(TT) > confdesemp.rango(2)) And (notas_desemp.porcentaje(TT) <= confdesemp.rango(1)) Then
                                                    Printer.Print "ALTO"
                                                End If
                                                If notas_desemp.porcentaje(TT) > confdesemp.rango(1) Then
                                                    Printer.Print "SUPERIOR"
                                                End If
                                                Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                            End If
                                        Next TT
                                        Close #NAR
                                        NAR = NAR - 1
                                    End If
                                End If
                            Wend
                            Close #NAR
                            NAR = NAR - 1
                        End If
                    End If
                
                End If
            Wend
            Close #NAR
            'NAR = NAR - 1
            Printer.Line (15.7, 5)-(15.7, Printer.CurrentY)
            Printer.Line (16.8, 5)-(16.8, Printer.CurrentY)
            Printer.Line (18.1, 5)-(18.1, Printer.CurrentY)
            Printer.Print ""
            Printer.FontSize = 9
            Printer.CurrentX = 1
            Printer.FontBold = True
            ' Abrir parámetros de promoción
            If (Dir(Ruta & "rangpro.txt") <> "") And (Dir(Ruta & "promovido.txt") <> "") Then
                Open Ruta & "rangpro.txt" For Input As #NAR
                Input #NAR, rus, fis
                Close #NAR
                Open Ruta & "promovido.txt" For Input As #NAR
                Input #NAR, SAPO2, SAPO3, SAPO4
                Close #NAR
            End If
            If MateriaX <= rus Then
                Printer.Print RTrim(SAPO2)
            End If
            If (MateriaX > rus) And (MateriaX <= fis) Then
                Printer.Print RTrim(SAPO3)
            End If
            If MateriaX > fis Then
                Printer.Print RTrim(SAPO4)
            End If
            Printer.FontBold = False
            Printer.Print ""
            If VeriNivela = True Then
                Printer.CurrentX = 1
                Printer.FontBold = True
                Printer.Print "NIVELACIÓN"
                Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                Printer.CurrentX = 1
                Printer.Print "ASIGNATURA";
                Printer.CurrentX = 11
                Printer.Print "VAL";
                Printer.CurrentX = 12.5
                Printer.Print "DESEMPEÑO";
                Printer.CurrentX = 15
                Printer.Print "FECHA / ACTA"
                Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                'Printer.Print NivelaTXT
                Printer.FontBold = False
                Cortar = Split(NivelaTXT, "%")
                For Recorrer = 0 To UBound(Cortar)
                    'Printer.CurrentX = 1
                    Cortar2 = Split(Cortar(Recorrer), "$")
                    For Recorrer2 = 0 To UBound(Cortar2)
                        If Recorrer2 = 0 Then
                            Printer.CurrentX = 1
                        End If
                        If Recorrer2 = 1 Then
                            Printer.CurrentX = 12.5
                            If (Cortar2(1)) <= confdesemp.rango(3) Then
                                Printer.Print "BAJO";
                            End If
                            If ((Cortar2(1)) > confdesemp.rango(3)) And ((Cortar2(1)) <= confdesemp.rango(2)) Then
                                Printer.Print "BÁSICO";
                            End If
                            If ((Cortar2(1)) > confdesemp.rango(2)) And ((Cortar2(1)) <= confdesemp.rango(1)) Then
                                Printer.Print "ALTO";
                            End If
                            If (Cortar2(1)) > confdesemp.rango(1) Then
                                Printer.Print "SUPERIOR";
                            End If
                            Printer.CurrentX = 11
                        End If
                        If Recorrer2 = 2 Then
                            Printer.CurrentX = 15
                        End If
                        If Recorrer2 <> 1 Then
                            Printer.Print Cortar2(Recorrer2);
                        Else
                            Printer.Print Format(Cortar2(Recorrer2), "#.00");
                        End If
                    Next Recorrer2
                    Printer.Print ""
                Next Recorrer
                Printer.Line (1, Printer.CurrentY - 0.2)-(21, Printer.CurrentY - 0.2)
                Printer.Print ""
            End If
            Printer.CurrentX = 1
            Printer.FontBold = True
            Printer.Print "OBSERVACIONES:"
            Printer.FontBold = False
            Printer.Print ""
            J = 0
            If (Dir(Ruta & "lrf" & Frame2.Caption & ".lrf") <> "") And (Dir(Ruta & "orf" & Frame2.Caption & ".orf") <> "") Then
                cona = 0
                Open Ruta & "lrf" & Frame2.Caption & ".lrf" For Random As #NAR Len = Len(leyfin)
                While Not EOF(NAR)
                    cona = cona + 1
                    Get #NAR, cona, leyfin
                    If Val(leyfin.num_carnet) = Val(alugru.num_carnet) Then
                        NAR = FreeFile
                        Open Ruta & "orf" & Frame2.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
                        For I = 1 To 5
                            If leyfin.fnob(I) <> 0 Then
                                Get #NAR, leyfin.fnob(I), obsfin
                                Printer.CurrentX = 1
                                
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
            'Printer.Print ""
            If Check1.Value = 1 Then
                'Firma del rector y secretaria académica
'                Printer.Line (3, Printer.CurrentY)-(8.5, Printer.CurrentY)
'                Printer.Line (13, Printer.CurrentY)-(18.5, Printer.CurrentY)
'                Printer.CurrentY = Printer.CurrentY + 0.1
'                Printer.CurrentX = 4
'                Printer.Print ini.Rector;
                Printer.Line (3, Printer.CurrentY)-(7.7, Printer.CurrentY)
                Printer.Line (13, Printer.CurrentY)-(18, Printer.CurrentY)
                Printer.CurrentY = Printer.CurrentY + 0.1
                Printer.CurrentX = 3.5
                Printer.Print ini.Rector;
                'Printer.CurrentX = 12.3
                Printer.CurrentX = 15.6 - ((Len(RTrim(ini.secretario)) / 4.8) / 2)
                Printer.Print ini.secretario
                'Printer.CurrentX = 5.2
                Printer.CurrentX = 4.5
                Printer.Print vini.VRector;
                Printer.CurrentX = 14
                Printer.Print "Secretaria Académica"
            Else
                'Firma del rector y director de grupo
                Printer.Line (3, Printer.CurrentY)-(7.7, Printer.CurrentY)
                Printer.Line (13, Printer.CurrentY)-(18, Printer.CurrentY)
                Printer.CurrentY = Printer.CurrentY + 0.1
                Printer.CurrentX = 3.5
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
            End If
            Screen.MousePointer = 0
            Printer.EndDoc
        Next VV
        Printer.PaperSize = 1
        Unload Me
        Screen.MousePointer = 0
    End If
End If



End Sub

Private Sub Form_Load()
'Option1.Value = True
Option3.Value = True
Text2.MaxLength = 3
Text3.MaxLength = 3
'Text1.Enabled = False
'Command1.Enabled = False
'Command2.Enabled = False
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
