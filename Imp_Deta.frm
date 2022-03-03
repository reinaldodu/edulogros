VERSION 5.00
Begin VB.Form Imp_Deta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión boletín detallado"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   Icon            =   "Imp_Deta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
      Begin VB.CheckBox Check1 
         Caption         =   "Periodo final"
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Có&digos"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Final..."
         Height          =   195
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inicial..."
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox Ver_Periodo 
         Height          =   315
         ItemData        =   "Imp_Deta.frx":0442
         Left            =   840
         List            =   "Imp_Deta.frx":0455
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox Ver_grupo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   480
      End
   End
End
Attribute VB_Name = "Imp_Deta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nar2 As Integer

Private Sub MODU3()
If (Printer.CurrentY > 23.5) Then
   Printer.Line (6.4, 3.8)-(6.4, Printer.CurrentY)
   Printer.NewPage
   'L = 1
   Nar2 = FreeFile
   Open Ruta & "inicial.edu" For Input As #Nar2
   Input #Nar2, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
   Close #Nar2
   Open Ruta & Ver_grupo.Text & ".gru" For Random As #Nar2 Len = Len(alugru)
   Get #Nar2, VV, alugru
   Close #Nar2
   Open Ruta & "prinalu.edu" For Random As #Nar2 Len = Len(alumno)
   Get #Nar2, (Val(alugru.num_carnet)), alumno
   Close #Nar2
    Printer.Font.Size = 14
    Printer.CurrentY = 0.5
    Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
    Printer.FontBold = True
    Printer.Print ini.nombre
    'Printer.CurrentX = 7
    Printer.Font.Size = 11
    'Printer.CurrentX = 10.2 - (Len("(CHIA - Aprobación oficial No.001437 septiembre 18 de 1997)") / 5.2) / 2
    'Printer.Print "(CHIA - Aprobación oficial No.001437 septiembre 18 de 1997)"
    Printer.CurrentX = 10.2 - (Len(ini.Rector) / 5.2) / 2
    Printer.Print ini.Rector
    Printer.Font.Size = 14
    Printer.CurrentX = 10.2 - ((Len("INFORME DE EVALUACION") / 3.3) / 2)
    Printer.Print "INFORME DE EVALUACION"
    Printer.FontBold = False
    Printer.Print ""
    Printer.Font.Size = 10
    Printer.CurrentX = 0.5
    Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
    Printer.CurrentX = 16.7
    Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
    Printer.CurrentX = 0.5
    Printer.Print "GRADO: " & RE22;
    Printer.CurrentX = 6
    Printer.Print "GRUPO: " & Ver_grupo.Text;
    Printer.CurrentX = 12.7
    Printer.Print "No.carnet: " & alumno.n_carnet;
    Printer.CurrentX = 16.7
    Printer.Print "PERIODO: " & Ver_Periodo.Text
    Printer.Print ""
    Printer.Font.Size = 12
    Printer.FontBold = True
    Printer.Line (0.2, 3.8)-(20.2, 3.8)
    Printer.CurrentX = 2.3
    Printer.Print "A R E A S";
    Printer.CurrentX = 10.5
    Printer.Print "O B S E R V A C I O N E S"
    Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
    Printer.FontBold = False
    Printer.CurrentY = 4.9
    Printer.Font.Size = 8
    YF = 4.9
End If
End Sub

Private Sub MODU2()
If (Printer.CurrentY > 26) Then
   Printer.Line (6.4, 3.8)-(6.4, Printer.CurrentY)
   Printer.NewPage
   'L = 1
   Nar2 = FreeFile
   Open Ruta & "inicial.edu" For Input As #Nar2
   Input #Nar2, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
   Close #Nar2
   Open Ruta & Ver_grupo.Text & ".gru" For Random As #Nar2 Len = Len(alugru)
   Get #Nar2, VV, alugru
   Close #Nar2
   Open Ruta & "prinalu.edu" For Random As #Nar2 Len = Len(alumno)
   Get #Nar2, (Val(alugru.num_carnet)), alumno
   Close #Nar2
    Printer.Font.Size = 14
    Printer.CurrentY = 0.5
    Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
    Printer.FontBold = True
    Printer.Print ini.nombre
    'Printer.CurrentX = 7
    Printer.Font.Size = 11
    'Printer.CurrentX = 10.2 - (Len("(CHIA - Aprobación oficial No.001437 septiembre 18 de 1997)") / 5.2) / 2
    'Printer.Print "(CHIA - Aprobación oficial No.001437 septiembre 18 de 1997)"
    Printer.CurrentX = 10.2 - (Len(ini.Rector) / 5.2) / 2
    Printer.Print ini.Rector
    Printer.Font.Size = 14
    Printer.CurrentX = 10.2 - ((Len("INFORME DE EVALUACION") / 3.3) / 2)
    Printer.Print "INFORME DE EVALUACION"
    Printer.FontBold = False
    Printer.Print ""
    Printer.Font.Size = 10
    Printer.CurrentX = 0.5
    Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
    Printer.CurrentX = 16.7
    Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
    Printer.CurrentX = 0.5
    Printer.Print "GRADO: " & RE22;
    Printer.CurrentX = 6
    Printer.Print "GRUPO: " & Ver_grupo.Text;
    Printer.CurrentX = 12.7
    Printer.Print "No.carnet: " & alumno.n_carnet;
    Printer.CurrentX = 16.7
    Printer.Print "PERIODO: " & Ver_Periodo.Text
    Printer.Print ""
    Printer.Font.Size = 12
    Printer.FontBold = True
    Printer.Line (0.2, 3.8)-(20.2, 3.8)
    Printer.CurrentX = 2.3
    Printer.Print "A R E A S";
    Printer.CurrentX = 10.5
    Printer.Print "O B S E R V A C I O N E S"
    Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
    Printer.FontBold = False
    Printer.CurrentY = 4.9
    Printer.Font.Size = 8
    YF = 4.9
End If
End Sub

Private Sub CUADRO()
Printer.Font.Size = 8
YC = Printer.CurrentY
Printer.Line (0.3, Printer.CurrentY)-(6, Printer.CurrentY)
Printer.Line (0.3, Printer.CurrentY + 0.1)-(6, Printer.CurrentY + 0.1)
Printer.CurrentX = 0.6
Printer.Font.Bold = True
Printer.Print "P";
Printer.CurrentX = 2.1
Printer.Print "EVALUACION";
Printer.CurrentX = 5.3
Printer.Print "FA"
Printer.Font.Bold = False
Printer.Line (0.3, Printer.CurrentY)-(6, Printer.CurrentY)
Printer.Line (0.3, Printer.CurrentY + 0.1)-(6, Printer.CurrentY + 0.1)
For J = 1 To 5
    If ((J <= lwe) Or ((J = 5) And Check1.Value = 1)) Then
        NV = " "
        If Dir(Ruta & Ver_grupo.Text & argra.num_area & J & ".obs") <> "" Then
            Nar2 = FreeFile
            z = 0
            Open Ruta & Ver_grupo.Text & argra.num_area & J & ".obs" For Random As #Nar2 Len = Len(notas)
            While Not EOF(Nar2)
                z = z + 1
                Get #Nar2, z, notas
                If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                    If RTrim(notas.FA) = "D" Then
                        NV = "DEFICIENTE"
                    End If
                    If RTrim(notas.FA) = "I" Then
                        NV = "INSUFICIENTE"
                    End If
                    If RTrim(notas.FA) = "A" Then
                        NV = "ACEPTABLE"
                    End If
                    If RTrim(notas.FA) = "S" Then
                        NV = "SOBRESALIENTE"
                    End If
                    If RTrim(notas.FA) = "E" Then
                        NV = "EXCELENTE"
                    End If
                    Printer.CurrentX = 0.6
                    If J <> 5 Then
                        Printer.Print J;
                    Else
                        Printer.Font.Bold = True
                        Printer.Print "F";
                    End If
                    Printer.CurrentX = 2.1
                    Printer.Print NV;
                    Printer.CurrentX = 5.3
                    Printer.Print notas.FA
                    Printer.Font.Bold = False
                    Printer.Line (0.3, Printer.CurrentY)-(6, Printer.CurrentY)
                    Close #Nar2
                    GoTo ter_Mat
                End If
            Wend
            Close #Nar2
        Else
            Printer.CurrentX = 0.6
            If J <> 5 Then
                Printer.Print J
            Else
                Printer.Font.Bold = True
                Printer.Print "F"
                Printer.Font.Bold = False
            End If
            Printer.Line (0.3, Printer.CurrentY)-(6, Printer.CurrentY)
        End If
    End If
ter_Mat:
Next J
Printer.Line (0.3, YC)-(0.3, Printer.CurrentY)
Printer.Line (1, YC)-(1, Printer.CurrentY)
Printer.Line (5, YC)-(5, Printer.CurrentY)
Printer.Line (6, YC)-(6, Printer.CurrentY)
Printer.Print ""
YF = Printer.CurrentY
End Sub

Private Sub Command1_Click()
If Dir(Ruta & Ver_grupo.Text & ".gru") = "" Then
    MsgBox "NO EXISTE ESTE GRUPO", 48, "IMPRIMIR BOLETINES"
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Ver_grupo.Text Then
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
If Ver_Periodo.Text = "PRIMERO" Then
    lwe = 1
End If
If Ver_Periodo.Text = "SEGUNDO" Then
    lwe = 2
End If
If Ver_Periodo.Text = "TERCERO" Then
    lwe = 3
End If
If Ver_Periodo.Text = "CUARTO" Then
    lwe = 4
End If
If Ver_Periodo.Text = "FINAL" Then
    lwe = 5
End If
ret = 0
Open Ruta & Ver_grupo.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
If Option1.Value = True Then
    s = 1
    q = ret - 1
    MS1 = "DESEA IMPRIMIR LOS BOLETINES DEL GRUPO " & Ver_grupo.Text & " DEL PERIODO " & Ver_Periodo.Text & "?"
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
    MS1 = "DESEA IMPRIMIR LOS BOLETINES DEL GRUPO " & Ver_grupo.Text & ", DESDE EL CODIGO " & Text1.Text & " HASTA EL CODIGO " & Text2.Text & " DEL PERIODO " & Ver_Periodo.Text & "?"
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR BOLETINES")
If RESP = vbYes Then
Screen.MousePointer = 11
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
Close #NAR
Printer.ScaleMode = 7
For VV = s To q
    'L = 1
    Open Ruta & Ver_grupo.Text & ".gru" For Random As #NAR Len = Len(alugru)
    Get #NAR, VV, alugru
    Close #NAR
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, (Val(alugru.num_carnet)), alumno
    Close #NAR
    Printer.Font.Size = 14
    Printer.CurrentY = 0.5
    Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
    Printer.FontBold = True
    Printer.Print ini.nombre
    'Printer.CurrentX = 7
    Printer.Font.Size = 11
    'Printer.CurrentX = 10.2 - (Len("(CHIA - Aprobación oficial No.001437 septiembre 18 de 1997)") / 5.2) / 2
    'Printer.Print "(CHIA - Aprobación oficial No.001437 septiembre 18 de 1997)"
    Printer.CurrentX = 10.2 - (Len(ini.Rector) / 5.2) / 2
    Printer.Print ini.Rector
    Printer.Font.Size = 14
    Printer.CurrentX = 10.2 - ((Len("INFORME DE EVALUACION") / 3.3) / 2)
    Printer.Print "INFORME DE EVALUACION"
    Printer.FontBold = False
    Printer.Print ""
    Printer.Font.Size = 10
    Printer.CurrentX = 0.5
    Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
    Printer.CurrentX = 16.7
    Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
    Printer.CurrentX = 0.5
    Printer.Print "GRADO: " & RE22;
    Printer.CurrentX = 6
    Printer.Print "GRUPO: " & Ver_grupo.Text;
    Printer.CurrentX = 12.7
    Printer.Print "No.carnet: " & alumno.n_carnet;
    Printer.CurrentX = 16.7
    Printer.Print "PERIODO: " & Ver_Periodo.Text
    Printer.Print ""
    Printer.Font.Size = 12
    Printer.FontBold = True
    Printer.Line (0.2, 3.8)-(20.2, 3.8)
    Printer.CurrentX = 2.3
    Printer.Print "A R E A S";
    Printer.CurrentX = 10.5
    Printer.Print "O B S E R V A C I O N E S"
    Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
    Printer.FontBold = False
    Printer.CurrentY = 4.9
    Printer.Font.Size = 8
    cona = 0
    YF = 4.9
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Ver_grupo.Text Then
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            '***ENCONTRAR EL CURRENTY MAS LARGO***
            If YF > Printer.CurrentY Then
                Printer.CurrentY = YF
            End If
            If Format(Printer.CurrentY, "##.#") <> 4.9 Then
                Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
                Printer.Print ""
            End If
            Call MODU3
            YI = Printer.CurrentY
            Printer.Font.Size = 10
            Printer.FontBold = True
            Printer.CurrentX = 0.3
            Printer.Print RTrim(mate.nom)
            Printer.Print ""
            Printer.Font.Size = 8
            Printer.FontBold = False
            '***FUNCION DE IMPRESION DEL CUADRO AQUI***
            Call CUADRO
            Printer.CurrentY = YI
            If Dir(Ruta & Ver_grupo.Text & argra.num_area & lwe & ".obs") <> "" Then
                z = 0
                Open Ruta & Ver_grupo.Text & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
                While Not EOF(NAR)
                    z = z + 1
                    Get #NAR, z, notas
                    If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                        NObs = False
                        For J = 1 To 4
                            h = 0
                            '***CONTADOR DE OBSERVACIONES***
                            CoB = 1
                            NH = 0
                            If J = 1 Then
                                VR = "F"
                                DVR = "FORTALEZAS:"
                            End If
                            If J = 2 Then
                                VR = "D"
                                DVR = "DIFICULTADES:"
                            End If
                            If J = 3 Then
                                VR = "S"
                                DVR = "RECOMENDACIONES:"
                            End If
                            If J = 4 Then
                                VR = ""
                                DVR = ""
                            End If
                            For I = 1 To 10
                                If notas.area(I) <> 0 Then
                                    NAR = FreeFile
                                    Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                                    Get #NAR, notas.area(I), logru
                                    Close #NAR
                                    If (logru.indicador = VR) And (NH = 0) Then
                                        If (J <> 1) And (NObs = True) Then
                                            Printer.Line (6.4, Printer.CurrentY)-(20.2, Printer.CurrentY)
                                            Printer.Print ""
                                            Call MODU2
                                        End If
                                        Printer.FontBold = True
                                        Printer.FontSize = 9
                                        Printer.CurrentX = 6.5
                                        Printer.Print DVR
                                        Call MODU2
                                        Printer.FontSize = 8
                                        Printer.FontBold = False
                                        NH = 1
                                        h = 1
                                        NObs = True
                                    End If
                                    If RTrim(logru.indicador) = VR Then
                                        X = 104
                                        L1 = Left(logru.observ, X)
                                        While Right(L1, 1) <> " "
                                            If X = 1 Then
                                                GoTo tolo
                                            End If
                                            X = X - 1
                                            L1 = Left(L1, X)
                                        Wend
tolo:
                                        Printer.CurrentX = 6.5
                                        Printer.Print CoB;
                                        Printer.Print ". ";
                                        CoB = CoB + 1
                                        Printer.CurrentX = 7
                                        Printer.Print L1
                                        'L = L + 1
                                        Call MODU2
                                        Y = Len(L1)
                                        Y = 200 - Y
                                        L2 = Right(logru.observ, Y)
                                        If RTrim(L2) <> "" Then
                                            Printer.CurrentX = 7
                                            Printer.Print L2
                                            'L = L + 1
                                            Call MODU2
                                        End If
                                    End If
                                    NAR = NAR - 1
                                End If
                            Next I
                        Next J
                        Printer.Print ""
                        'L = L + 1
                        Call MODU2
                    End If
                Wend
                Close #NAR
            End If
            NAR = NAR - 1
        End If
    Wend
    'Printer.TextHeight
    Close #NAR
    If YF > Printer.CurrentY Then
        Printer.CurrentY = YF
    End If
    Printer.Line (6.4, 3.8)-(6.4, Printer.CurrentY)
    Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
    If (Printer.CurrentY >= 17) Then
        Printer.NewPage
        Printer.Font.Size = 14
        Printer.CurrentY = 0.5
        Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
        Printer.FontBold = True
        Printer.Print ini.nombre
        'Printer.CurrentX = 7
        Printer.Font.Size = 11
        'Printer.CurrentX = 10.2 - (Len("(CHIA - Aprobación oficial No.001437 septiembre 18 de 1997)") / 5.2) / 2
        'Printer.Print "(CHIA - Aprobación oficial No.001437 septiembre 18 de 1997)"
        Printer.CurrentX = 10.2 - (Len(ini.Rector) / 5.2) / 2
        Printer.Print ini.Rector
        Printer.Font.Size = 14
        Printer.CurrentX = 10.2 - ((Len("INFORME DE EVALUACION") / 3.3) / 2)
        Printer.Print "INFORME DE EVALUACION"
        Printer.FontBold = False
        Printer.Print ""
        Printer.Font.Size = 10
        Printer.CurrentX = 0.5
        Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
        Printer.CurrentX = 16.7
        Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
        Printer.CurrentX = 0.5
        Printer.Print "GRADO: " & RE22;
        Printer.CurrentX = 6
        Printer.Print "GRUPO: " & Ver_grupo.Text;
        Printer.CurrentX = 12.7
        Printer.Print "No.carnet: " & alumno.n_carnet;
        Printer.CurrentX = 16.7
        Printer.Print "PERIODO: " & Ver_Periodo.Text
        Printer.Print ""
        Printer.Font.Size = 12
        Printer.FontBold = True
        Printer.Line (0.2, 3.8)-(20.2, 3.8)
        Printer.CurrentX = 2.3
        Printer.Print "A R E A S";
        Printer.CurrentX = 10.5
        Printer.Print "O B S E R V A C I O N E S"
        Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
        Printer.FontBold = False
        Printer.CurrentY = 4.9
        Printer.Font.Size = 8
        YF = 4.9
    Else
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
    End If
    Open Ruta & "leyenda.edu" For Input As #NAR
    Input #NAR, leye.ly1, leye.ly2, leye.ly3, leye.ly4, leye.ly5, leye.ly6, leye.ly7
    Close #NAR
    Printer.Font.Size = 10
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
    Printer.CurrentX = 0.5
    Printer.Print leye.ly7
    Printer.Print ""
    Printer.Print ""
    Printer.Line (0.5, Printer.CurrentY)-(19.7, Printer.CurrentY)
    Printer.Print ""
    Printer.Print ""
    Printer.Line (0.5, Printer.CurrentY)-(19.7, Printer.CurrentY)
    Printer.Print ""
    Printer.Print ""
    Printer.Line (0.5, Printer.CurrentY)-(19.7, Printer.CurrentY)
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Line (2.5, Printer.CurrentY)-(7.5, Printer.CurrentY)
    Printer.Line (13, Printer.CurrentY)-(18, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.CurrentX = 3.5
    'Printer.Print "Firma del Directora.";
    Printer.Print Rector;
    Printer.CurrentX = 15.6 - ((Len(PERI) / 4.8) / 2)
    Printer.Print PERI
    
    Printer.CurrentX = 4.5
    Printer.Print "Directora.";
    
    Printer.CurrentX = 14
    Printer.Print "Directora de grupo."
    Printer.NewPage
Next VV
Printer.EndDoc
Unload Me
Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Load()
Option1.Value = True
Text1.MaxLength = 2
Text2.MaxLength = 2
Ver_Periodo.Text = Ver_Periodo.List(0)
If Dir(Ruta & "infcur.edu") <> "" Then
    Command1.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Ver_grupo.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Ver_grupo.Text = Ver_grupo.List(0)
Else
    Command1.Enabled = False
End If
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
