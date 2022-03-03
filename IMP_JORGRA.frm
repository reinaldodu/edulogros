VERSION 5.00
Begin VB.Form IMP_JORGRA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de boletines por jornada y grado"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "IMP_JORGRA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   320
      Left            =   3765
      TabIndex        =   14
      Top             =   420
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I&mprimir"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "IMP_JORGRA.frx":0442
         Left            =   1440
         List            =   "IMP_JORGRA.frx":0455
         TabIndex        =   1
         Text            =   "PRIMERO"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "IMP_JORGRA.frx":0483
         Left            =   1440
         List            =   "IMP_JORGRA.frx":04B7
         TabIndex        =   3
         Text            =   "PREKINDER"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "IMP_JORGRA.frx":0540
         Left            =   1440
         List            =   "IMP_JORGRA.frx":0550
         TabIndex        =   2
         Text            =   "UNICA"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opciones"
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   3855
         Begin VB.TextBox Txt_Espa 
            Height          =   320
            Left            =   2955
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0"
            Top             =   300
            Width           =   390
         End
         Begin VB.CommandButton Command2 
            Caption         =   "+"
            Height          =   320
            Left            =   3330
            TabIndex        =   13
            Top             =   300
            Width           =   195
         End
         Begin VB.CheckBox Check2 
            Caption         =   "&Imprimir toda la jornada"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1680
            TabIndex        =   8
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Encabezado"
            Height          =   255
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Sin Formato"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Con Formato"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&PERIODO :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GRADO    :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   1920
         Width           =   810
      End
   End
End
Attribute VB_Name = "IMP_JORGRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
    Txt_Espa.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
Else
    Txt_Espa.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
End If
End Sub

Private Sub Combo1_Change()
If Combo1.Text <> Combo1.List(0) Then
    Combo1.Text = Combo1.List(0)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim notas As notis
'Dim argra As areagr
'Dim logru As logris
'Dim mate As infomater
'Dim alugru As grupoalu
'Dim ini As inicio
'Dim leye As leyendis
'Dim icur As inforcur
'Dim profe As maestropro
PERI = ""
If Check2.Value = 1 Then
    MS1 = "Desea imprimir todos los boletines de la jornada " & Format(Combo1.Text, "<") & "?"
Else
    MS1 = "Desea imprimir todos los boletines del grado " & Format(Combo2.Text, "<") & ", jornada " & Format(Combo1.Text, "<") & "?"
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir boletines")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        TTT = ""
        If Check2.Value = 1 Then
            If RTrim(icur.jornada) = Combo1.Text Then
                TTT = RTrim(icur.nom)
                RE22 = RTrim(icur.grado)
                JOJI = RTrim(icur.jornada)
                SP = RTrim(icur.director)
            End If
        Else
            If (RTrim(icur.jornada) = Combo1.Text) And (RTrim(icur.grado) = Combo2.Text) Then
                TTT = RTrim(icur.nom)
                RE22 = RTrim(icur.grado)
                JOJI = RTrim(icur.jornada)
                SP = RTrim(icur.director)
            End If
        End If
        If TTT <> "" Then
            NAR = FreeFile
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
            Open Ruta & TTT & ".gru" For Random As #NAR Len = Len(alugru)
            While Not EOF(NAR)
                ret = ret + 1
                Get #NAR, ret, alugru
            Wend
            Close #NAR
            NAR = NAR - 1
        Else
            GoTo impregran
        End If
        q = ret - 1
        NAR = FreeFile
        Open Ruta & "inicial.edu" For Input As #NAR
        Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
        Close #NAR
        Printer.ScaleMode = 7
        For VV = 1 To q
            L = 0
            Open Ruta & TTT & ".gru" For Random As #NAR Len = Len(alugru)
            Get #NAR, VV, alugru
            Close #NAR
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, (Val(alugru.num_carnet)), alumno
            Close #NAR
            If Option2.Value = True Then
               Printer.Font.Size = 14
               Printer.CurrentY = 1
               Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
               Printer.FontBold = True
               Printer.Print ini.nombre
               Printer.CurrentX = 7.4
               Printer.Print "INFORME DESCRIPTIVO"
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
               Printer.Print "GRUPO: " & TTT;
               Printer.CurrentX = 12.7
               Printer.Print "No.carnet: " & alumno.n_carnet;
               Printer.CurrentX = 16.7
               Printer.Print "PERIODO: " & Combo3.Text
               Printer.Print ""
               Printer.Font.Size = 12
               Printer.CurrentX = 0.5
               Printer.Print "A R E A S";
               Printer.CurrentX = 5
               Printer.Print "I.H";
               Printer.CurrentX = 5.7
               Printer.Print "FA";
               Printer.CurrentX = 6.5
               Printer.Print "J.V";
               Printer.CurrentX = 7.3
               Printer.Print "IND";
               Printer.CurrentX = 8.2
               Printer.Print "O B S E R V A C I O N E S"
               Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
               Printer.CurrentY = 4.9
            Else
               If Check1.Value = 1 Then
                  Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.9
               Else
                  Printer.CurrentY = 2.9
               End If
               Printer.Font.Size = 10
               Printer.CurrentX = 3.2
               Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
               Printer.CurrentX = 17.8
               Printer.Print Format(Date, "mmm/dd/yyyy")
               If Check1.Value = 1 Then
                  Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 3.4
               Else
                  Printer.CurrentY = 3.4
               End If
               Printer.CurrentX = 2.3
               Printer.Print RE22;
               Printer.CurrentX = 7.5
               Printer.Print TTT;
               Printer.CurrentX = 13.5
               Printer.Print alumno.n_carnet;
               Printer.CurrentX = 18.3
               Printer.Print Combo3.Text
               Printer.CurrentY = 4.9
            End If
            Printer.Font.Size = 8
            cona = 0
            Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
            While Not EOF(NAR)
                cona = cona + 1
                Get #NAR, cona, argra
                If RTrim(argra.nom_grup) = TTT Then
                    If Dir(Ruta & TTT & argra.num_area & lwe & ".obs") <> "" Then
                        NAR = FreeFile
                        z = 0
                        Open Ruta & TTT & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
                        While Not EOF(NAR)
                            z = z + 1
                            Get #NAR, z, notas
                            If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                                NAR = FreeFile
                                Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                                Get #NAR, argra.num_area, mate
                                Close #NAR
                                Printer.CurrentX = 0.3
                                Printer.Print RTrim(mate.nom) & " " & "(" & mate.num & ")";
                                Printer.CurrentX = 5.1
                                Printer.Print argra.ih;
                                Printer.CurrentX = 5.8
                                Printer.Print notas.FA;
                                Printer.CurrentX = 6.6
                                Printer.Print notas.FA;
                                h = 0
                                For I = 1 To 10
                                    If notas.area(I) <> 0 Then
                                        h = 1
                                        Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                                        Get #NAR, notas.area(I), logru
                                        Close #NAR
                                        Printer.CurrentX = 7.5
                                        Printer.Print logru.indicador;
                                        Printer.CurrentX = 8.2
                                        X = 92
                                        L1 = Left(logru.observ, X)
                                        While Right(L1, 1) <> " "
                                            If X = 1 Then
                                                GoTo tolo
                                            End If
                                            X = X - 1
                                            L1 = Left(L1, X)
                                        Wend
tolo:
                                        Printer.CurrentX = 8.2
                                        Printer.Print L1
                                        L = L + 1
                                        Call MODULAR
                                        Y = Len(L1)
                                        Y = 200 - Y
                                        L2 = Right(logru.observ, Y)
                                        If RTrim(L2) <> "" Then
                                            Printer.CurrentX = 8.2
                                            Printer.Print L2
                                            L = L + 1
                                            Call MODULAR
                                        End If
                                    End If
                                Next I
                                If h = 0 Then
                                   Printer.Print ""
                                   L = L + 1
                                   Call MODULAR
                                End If
                                Printer.Print ""
                                L = L + 1
                                Call MODULAR
                                Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
                                Printer.Print ""
                                L = L + 1
                                Call MODULAR
                                NAR = NAR - 1
                            End If
                        Wend
                        Close #NAR
                        NAR = NAR - 1
                    End If
                End If
            Wend
            Close #NAR
            Printer.CurrentY = Printer.CurrentY - 0.35
            Printer.Line (4.8, 4.2)-(4.8, Printer.CurrentY)
            Printer.Line (5.6, 4.2)-(5.6, Printer.CurrentY)
            Printer.Line (6.4, 4.2)-(6.4, Printer.CurrentY)
            Printer.Line (7.2, 4.2)-(7.2, Printer.CurrentY)
            Printer.Line (8, 4.2)-(8, Printer.CurrentY)
            If (62 - L) < 22 Then
               Printer.NewPage
               If Option1.Value = True Then
                  If Check1.Value = 1 Then
                     Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.9
                  Else
                     Printer.CurrentY = 2.9
                  End If
                  Printer.Font.Size = 10
                  Printer.CurrentX = 3.2
                  Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
                  Printer.CurrentX = 17.8
                  Printer.Print Format(Date, "mmm/dd/yyyy")
                  If Check1.Value = 1 Then
                     Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 3.4
                  Else
                     Printer.CurrentY = 3.4
                  End If
                  Printer.CurrentX = 2.3
                  Printer.Print RE22;
                  Printer.CurrentX = 7.5
                  Printer.Print TTT;
                  Printer.CurrentX = 13.5
                  Printer.Print alumno.n_carnet;
                  Printer.CurrentX = 18.3
                  Printer.Print Combo3.Text
                  Printer.CurrentY = 4.9
               Else
                  Printer.Font.Size = 14
                  Printer.CurrentY = 1
                  Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
                  Printer.FontBold = True
                  Printer.Print ini.nombre
                  Printer.CurrentX = 7.4
                  Printer.Print "INFORME DESCRIPTIVO"
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
                  Printer.Print "GRUPO: " & TTT;
                  Printer.CurrentX = 12.7
                  Printer.Print "No.carnet: " & alumno.n_carnet;
                  Printer.CurrentX = 16.7
                  Printer.Print "PERIODO: " & Combo3.Text
                  Printer.Print ""
                  Printer.Font.Size = 12
                  Printer.CurrentX = 0.5
                  Printer.Print "A R E A S";
                  Printer.CurrentX = 5
                  Printer.Print "I.H";
                  Printer.CurrentX = 5.7
                  Printer.Print "FA";
                  Printer.CurrentX = 6.5
                  Printer.Print "J.V";
                  Printer.CurrentX = 7.3
                  Printer.Print "IND";
                  Printer.CurrentX = 8.2
                  Printer.Print "O B S E R V A C I O N E S"
                  Printer.Font.Size = 8
                  Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
                  Printer.CurrentY = 4.9
               End If
            Else
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
        NAR = NAR - 1
        Printer.NewPage
impregran:
    Wend
    Close #NAR
    If PERI = "" Then
        MsgBox "No existe información para imprimir", 64, "Imprimir boletines"
    Else
        Printer.EndDoc
        Unload Me
    End If
End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Txt_Espa.Text = Txt_Espa.Text + 1
End Sub

Private Sub Command3_Click()
Txt_Espa.Text = Txt_Espa.Text - 1
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Impresión de boletines de toda una jornada o de un grado en específico."
End Sub

Private Sub Form_Load()
If Dir(Ruta & "infcur.edu") = "" Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If
Option1.Value = True
Txt_Espa.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub
Private Sub MODULAR()
'Dim alumno As maestroalum
'Dim alugru As grupoalu
'Dim ini As inicio
If (L Mod 62) = 0 Then
               Printer.Line (4.8, 4.2)-(4.8, Printer.CurrentY)
               Printer.Line (5.6, 4.2)-(5.6, Printer.CurrentY)
               Printer.Line (6.4, 4.2)-(6.4, Printer.CurrentY)
               Printer.Line (7.2, 4.2)-(7.2, Printer.CurrentY)
               Printer.Line (8, 4.2)-(8, Printer.CurrentY)
               Printer.NewPage
               L = 0
               NAR = FreeFile
               Open Ruta & "inicial.edu" For Input As #NAR
               Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
               Close #NAR
               Open Ruta & TTT & ".gru" For Random As #NAR Len = Len(alugru)
               Get #NAR, VV, alugru
               Close #NAR
               Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
               Get #NAR, (Val(alugru.num_carnet)), alumno
               Close #NAR
               If Option1.Value = True Then
                  If Check1.Value = 1 Then
                     Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.9
                  Else
                     Printer.CurrentY = 2.9
                  End If
                  Printer.Font.Size = 10
                  Printer.CurrentX = 3.2
                  Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
                  Printer.CurrentX = 17.8
                  Printer.Print Format(Date, "mmm/dd/yyyy")
                  If Check1.Value = 1 Then
                     Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 3.4
                  Else
                     Printer.CurrentY = 3.4
                  End If
                  Printer.CurrentX = 2.3
                  Printer.Print RE22;
                  Printer.CurrentX = 7.5
                  Printer.Print TTT;
                  Printer.CurrentX = 13.5
                  Printer.Print alumno.n_carnet;
                  Printer.CurrentX = 18.3
                  Printer.Print Combo3.Text
                  Printer.CurrentY = 4.9
               Else
                  Printer.Font.Size = 14
                  Printer.CurrentY = 1
                  Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
                  Printer.FontBold = True
                  Printer.Print ini.nombre
                  Printer.CurrentX = 7.4
                  Printer.Print "INFORME DESCRIPTIVO"
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
                  Printer.Print "GRUPO: " & TTT;
                  Printer.CurrentX = 12.7
                  Printer.Print "No.carnet: " & alumno.n_carnet;
                  Printer.CurrentX = 16.7
                  Printer.Print "PERIODO: " & Combo3.Text
                  Printer.Print ""
                  Printer.Font.Size = 12
                  Printer.CurrentX = 0.5
                  Printer.Print "A R E A S";
                  Printer.CurrentX = 5
                  Printer.Print "I.H";
                  Printer.CurrentX = 5.7
                  Printer.Print "FA";
                  Printer.CurrentX = 6.5
                  Printer.Print "J.V";
                  Printer.CurrentX = 7.3
                  Printer.Print "IND";
                  Printer.CurrentX = 8.2
                  Printer.Print "O B S E R V A C I O N E S"
                  Printer.Font.Size = 8
                  Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
                  Printer.CurrentY = 4.9
               End If
               Printer.Font.Size = 8
End If
End Sub
