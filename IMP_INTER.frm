VERSION 5.00
Begin VB.Form IMP_INTER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervalo de impresión"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "IMP_INTER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command3 
         Caption         =   "-"
         Height          =   320
         Left            =   2820
         TabIndex        =   14
         Top             =   630
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   320
         Left            =   2625
         TabIndex        =   13
         Top             =   630
         Width           =   195
      End
      Begin VB.TextBox Txt_Espa 
         Height          =   320
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   630
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Encabezado"
         Height          =   255
         Left            =   2265
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "&Sin Formato"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&Con Formato"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   405
         Left            =   1695
         Picture         =   "IMP_INTER.frx":0442
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Có&digos"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Final..."
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicial..."
         Height          =   195
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Width           =   540
      End
   End
End
Attribute VB_Name = "IMP_INTER"
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

Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim notas As notis
'Dim argra As areagr
'Dim logru As logris
'Dim mate As infomater
'Dim alugru As grupoalu
'Dim ini As inicio
'Dim leye As leyendis
If Option1.Value = True Then
s = 1
q = ret - 1
MS1 = "DESEA IMPRIMIR LOS BOLETINES DEL GRUPO " & Frame1.Caption & " DEL PERIODO " & CONS_NOTA.Combo3.Text & "?"
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
MS1 = "DESEA IMPRIMIR LOS BOLETINES DEL GRUPO " & Frame1.Caption & ", DESDE EL CODIGO " & Text1.Text & " HASTA EL CODIGO " & Text2.Text & " DEL PERIODO " & CONS_NOTA.Combo3.Text & "?"
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR BOLETINES")
If RESP = vbYes Then
Screen.MousePointer = 11
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
Printer.ScaleMode = 7
For VV = s To q
    L = 0
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
       Printer.CurrentX = 7.4
       Printer.Print "INFORME DESCRIPTIVO"
       Printer.FontBold = False
       Printer.Print ""
       Printer.Font.Size = 10
       Printer.CurrentX = 0.5
       Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
       'Printer.CurrentX = 12.7
       'Printer.Print "CIUDAD: " & ini.ciudad;
       Printer.CurrentX = 16.7
       Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
       Printer.CurrentX = 0.5
       Printer.Print "GRADO: " & RE22;
       Printer.CurrentX = 6
       Printer.Print "GRUPO: " & Frame1.Caption;
       Printer.CurrentX = 12.7
       Printer.Print "No.carnet: " & alumno.n_carnet;
       Printer.CurrentX = 16.7
       Printer.Print "PERIODO: " & CONS_NOTA.Combo3.Text
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
       'Printer.CurrentX = 13.3
       'Printer.Print ini.ciudad;
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
       Printer.Print Frame1.Caption;
       Printer.CurrentX = 13.5
       Printer.Print alumno.n_carnet;
       Printer.CurrentX = 18.3
       Printer.Print CONS_NOTA.Combo3.Text
       Printer.CurrentY = 4.9
    End If
    Printer.Font.Size = 8
    cona = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Frame1.Caption Then
            If Dir(Ruta & Frame1.Caption & argra.num_area & lwe & ".obs") <> "" Then
                NAR = FreeFile
                z = 0
                Open Ruta & Frame1.Caption & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
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
                                Call MODU
                                Y = Len(L1)
                                Y = 200 - Y
                                L2 = Right(logru.observ, Y)
                                If RTrim(L2) <> "" Then
                                    Printer.CurrentX = 8.2
                                    Printer.Print L2
                                    L = L + 1
                                    Call MODU
                                End If
                            End If
                        Next I
                        If h = 0 Then
                           Printer.Print ""
                           L = L + 1
                           Call MODU
                        End If
                        Printer.Print ""
                        L = L + 1
                        Call MODU
                        Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
                        Printer.Print ""
                        L = L + 1
                        Call MODU
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
       If Option3.Value = True Then
          If Check1.Value = 1 Then
             Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.9
          Else
             Printer.CurrentY = 2.9
          End If
          Printer.Font.Size = 10
          Printer.CurrentX = 3.2
          Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
          'Printer.CurrentX = 13.3
          'Printer.Print ini.ciudad;
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
          Printer.Print Frame1.Caption;
          Printer.CurrentX = 13.5
          Printer.Print alumno.n_carnet;
          Printer.CurrentX = 18.3
          Printer.Print CONS_NOTA.Combo3.Text
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
          'Printer.CurrentX = 12.7
          'Printer.Print "CIUDAD: " & ini.ciudad;
          Printer.CurrentX = 16.7
          Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
          Printer.CurrentX = 0.5
          Printer.Print "GRADO: " & RE22;
          Printer.CurrentX = 6
          Printer.Print "GRUPO: " & Frame1.Caption;
          Printer.CurrentX = 12.7
          Printer.Print "No.carnet: " & alumno.n_carnet;
          Printer.CurrentX = 16.7
          Printer.Print "PERIODO: " & CONS_NOTA.Combo3.Text
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
    Printer.Print ini.Rector;
    'Printer.Print Rector;
    Printer.CurrentX = 15.6 - ((Len(PERI) / 4.8) / 2)
    Printer.Print PERI
    
    Printer.CurrentX = 4.5
    Printer.Print vini.VRector;
    
    Printer.CurrentX = 14
    Printer.Print vini.VDirector
    Printer.NewPage
Next VV
Printer.EndDoc
Unload Me
Screen.MousePointer = 0
End If
End Sub

Private Sub Command2_Click()
Txt_Espa.Text = Txt_Espa.Text + 1
End Sub

Private Sub Command3_Click()
Txt_Espa.Text = Txt_Espa.Text - 1
End Sub

Private Sub Form_Load()
Option1.Value = True
Option3.Value = True
Text1.MaxLength = 2
Text2.MaxLength = 2
Frame1.Caption = CONS_NOTA.Combo4.Text
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
Private Sub MODU()
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
               Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
               Close #NAR
               Open Ruta & Frame1.Caption & ".gru" For Random As #NAR Len = Len(alugru)
               Get #NAR, VV, alugru
               Close #NAR
               Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
               Get #NAR, (Val(alugru.num_carnet)), alumno
               Close #NAR
               If Option3.Value = True Then
                  If Check1.Value = 1 Then
                     Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.9
                  Else
                     Printer.CurrentY = 2.9
                  End If
                  Printer.Font.Size = 10
                  Printer.CurrentX = 3.2
                  Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
                  'Printer.CurrentX = 13.3
                  'Printer.Print ini.ciudad;
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
                  Printer.Print Frame1.Caption;
                  Printer.CurrentX = 13.5
                  Printer.Print alumno.n_carnet;
                  Printer.CurrentX = 18.3
                  Printer.Print CONS_NOTA.Combo3.Text
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
                  'Printer.CurrentX = 12.7
                  'Printer.Print "CIUDAD: " & ini.ciudad;
                  Printer.CurrentX = 16.7
                  Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
                  Printer.CurrentX = 0.5
                  Printer.Print "GRADO: " & RE22;
                  Printer.CurrentX = 6
                  Printer.Print "GRUPO: " & Frame1.Caption;
                  Printer.CurrentX = 12.7
                  Printer.Print "No.carnet: " & alumno.n_carnet;
                  Printer.CurrentX = 16.7
                  Printer.Print "PERIODO: " & CONS_NOTA.Combo3.Text
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

