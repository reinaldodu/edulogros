VERSION 5.00
Begin VB.Form IMP_RECU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de recuperaciones"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "IMP_RECU.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "SEGUNDO"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   19
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seleccione periodo"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   3615
      Begin VB.CheckBox Check2 
         Caption         =   "CUARTO"
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "TERCERO"
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "PRIMERO"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   1920
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
         Picture         =   "IMP_RECU.frx":0442
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   3240
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
   Begin VB.Label Label3 
      Caption         =   "GRUPO:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "IMP_RECU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo1.SetFocus
End If
End Sub

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
If ((Check2(0).Value = 0) And (Check2(1).Value = 0) And (Check2(2).Value = 0) And (Check2(3).Value = 0)) Then
    MsgBox "SELECCIONE UN PERIODO ACADEMICO", 48, "ADVERTENCIA"
    Exit Sub
End If

ret = 0
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
If Option1.Value = True Then
    s = 1
    q = ret - 1
    MS1 = "DESEA IMPRIMIR LOS REPORTES DEL GRUPO " & Combo1.Text & "?"
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
    MS1 = "DESEA IMPRIMIR LOS REPORTES DEL GRUPO " & Combo1.Text & "?"
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR REPORTES")
If RESP = vbYes Then
Screen.MousePointer = 11
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR

Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If (Combo1.Text = RTrim(icur.nom)) Then
        seri = "1" & Left(icur.grado, 3)
    End If
Wend
Close #NAR

'seri = Left(Combo1.Text, 4)
'Error 1sexto-b
'seri = "1SEX"
Printer.ScaleMode = 7
For VV = s To q
    L = 1
    Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
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
       Printer.CurrentX = 6.2
       Printer.Print "REPORTE DE RECUPERACIONES"
       Printer.FontBold = False
       Printer.Print ""
       Printer.Font.Size = 10
       Printer.CurrentX = 0.5
       Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
       'Printer.CurrentX = 12.7
       'Printer.Print "CIUDAD: " & ini.ciudad;
       Printer.CurrentX = 16.7
       Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
       'Printer.CurrentX = 0.5
       'Printer.Print "GRADO: " & RE22;
       'Printer.CurrentX = 6
       Printer.CurrentX = 0.5
       Printer.Print "GRUPO: " & Combo1.Text;
       Printer.CurrentX = 12.7
       Printer.Print "No.carnet: " & alumno.n_carnet;
       'Printer.CurrentX = 16.7
       'Printer.Print "PERIODO: " & CONS_NOTA.Combo3.Text
       Printer.Print ""
       Printer.Print ""
       Printer.Font.Size = 12
       Printer.Print ""
       Printer.CurrentX = 0.5
       Printer.Print "A R E A S";
       Printer.CurrentX = 6
       Printer.Print "O B S E R V A C I O N E S"
       Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
       Printer.CurrentY = 4.9
    Else
       If Check1.Value = 1 Then
          Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.9
       Else
          Printer.CurrentY = 2.9
       End If
       Printer.Font.Size = 12
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
       Printer.Print Combo1.Text;
       Printer.CurrentX = 13.5
       Printer.Print alumno.n_carnet
       'Printer.Print CONS_NOTA.Combo3.Text
       Printer.CurrentY = 4.9
    End If
    For REC = 0 To 3
        If (Check2(REC).Value = 1) Then
            lwe = REC + 1
            If (lwe = 1) Then
                info = "PRIMER"
            End If
            If (lwe = 2) Then
                info = "SEGUNDO"
            End If
            If (lwe = 3) Then
                info = "TERCER"
            End If
            If (lwe = 4) Then
                info = "CUARTO"
            End If
        'End If
    Printer.Font.Size = 10
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.CurrentX = 0.3
    Printer.Print info & " PERIODO:"
    Printer.Print ""
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.Font.Size = 8
    cona = 0
    'NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #1 Len = Len(argra)
    While Not EOF(1)
        cona = cona + 1
        Get #1, cona, argra
        
        ver = 0
                
        
        If RTrim(argra.nom_grup) = Combo1.Text Then
          
            If Dir(Ruta & Combo1.Text & argra.num_area & lwe & ".obs") <> "" Then
                'NAR = FreeFile
                z = 0
                Open Ruta & Combo1.Text & argra.num_area & lwe & ".obs" For Random As #2 Len = Len(notas)
                While Not EOF(2)
                    z = z + 1
                    Get #2, z, notas
                                        
                    If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                    encuentra = 0
                    'RUTINA PARA BUSCAR MATERIAS CON R-D
                    For I = 1 To 10
                            If notas.area(I) <> 0 Then
                                h = 1
                                'NAR = FreeFile
                                Open Ruta & seri & argra.num_area & lwe & ".lgr" For Random As #4 Len = Len(logru)
                                Get #4, notas.area(I), logru
                                Close #4
                                'Printer.CurrentX = 7.5
                            If ((logru.indicador = "R") Or (logru.indicador = "D")) Then
                                encuentra = 1
                            End If
                            End If
                    Next I
                    If (encuentra = 1) Then
                    
                      If (ver = 0) Then
                        'NAR = FreeFile
                        Open Ruta & "materia.edu" For Random As #3 Len = Len(mate)
                        Get #3, argra.num_area, mate
                        Close #3
                        Printer.CurrentX = 0.3
                        Printer.Print RTrim(mate.nom) & " " & "(" & mate.num & ")";
                        ver = 1
                       End If
                        'Printer.CurrentX = 5.1
                        'Printer.Print argra.ih;
                        'Printer.CurrentX = 5.8
                        'printer.Print notas.FA;
                        'Printer.CurrentX = 6.6
                        'Printer.Print notas.FA;
                                              
                                                
                        h = 0
                        For I = 1 To 10
                            If notas.area(I) <> 0 Then
                                h = 1
                                'NAR = FreeFile
                                Open Ruta & seri & argra.num_area & lwe & ".lgr" For Random As #4 Len = Len(logru)
                                Get #4, notas.area(I), logru
                                Close #4
                                'Printer.CurrentX = 7.5
                            If ((logru.indicador = "R") Or (logru.indicador = "D")) Then
                                Printer.CurrentX = 5.1
                                Printer.Print logru.indicador;
                                'Printer.CurrentX = 8.2
                                Printer.CurrentX = 5.8
                                X = 105
                                L1 = Left(logru.observ, X)
                                While Right(L1, 1) <> " "
                                    If X = 1 Then
                                        GoTo tolo
                                    End If
                                    X = X - 1
                                    L1 = Left(L1, X)
                                Wend
tolo:
                                Printer.CurrentX = 5.8
                                Printer.Print L1
                                L = L + 1
                                Call MODU
                                Y = Len(L1)
                                Y = 200 - Y
                                L2 = Right(logru.observ, Y)
                                If RTrim(L2) <> "" Then
                                    Printer.CurrentX = 5.8
                                    Printer.Print L2
                                    L = L + 1
                                    Call MODU
                                End If
                            End If
                            End If
                        Next I
                    'Next REC
                        
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
                        'NAR = NAR - 1
                    End If
                    End If
                Wend
                Close #2
                'NAR = NAR - 1
            End If
            
        End If
        'Next rec
    Wend
    Close #1
    'Next REC
    Printer.CurrentY = Printer.CurrentY - 0.35
    Printer.Line (4.8, 4.2)-(4.8, Printer.CurrentY)
    Printer.Line (5.6, 4.2)-(5.6, Printer.CurrentY)
    'Printer.Line (6.4, 4.2)-(6.4, Printer.CurrentY)
    'printer.Line (7.2, 4.2)-(7.2, Printer.CurrentY)
    'Printer.Line (8, 4.2)-(8, Printer.CurrentY)
    'If (62 - L) < 14 Then
    'If (L Mod 60) = 0 Then
    If (L > 60) Then
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
          Printer.Print Combo1.Text;
          Printer.CurrentX = 13.5
          Printer.Print alumno.n_carnet;
          Printer.CurrentX = 18.3
          'Printer.Print CONS_NOTA.Combo3.Text
          Printer.CurrentY = 4.9
       Else
          Printer.Font.Size = 14
          Printer.CurrentY = 1
          Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
          Printer.FontBold = True
          Printer.Print ini.nombre
          Printer.CurrentX = 6.2
          Printer.Print "REPORTE DE RECUPERACIONES"
          Printer.FontBold = False
          Printer.Print ""
          Printer.Font.Size = 10
          Printer.CurrentX = 0.5
          Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
          'Printer.CurrentX = 12.7
          'Printer.Print "CIUDAD: " & ini.ciudad;
          Printer.CurrentX = 16.7
          Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
          'Printer.CurrentX = 0.5
          'Printer.Print "GRADO: " & RE22;
          'Printer.CurrentX = 6
          Printer.CurrentX = 0.5
          Printer.Print "GRUPO: " & Combo1.Text;
          Printer.CurrentX = 12.7
          Printer.Print "No.carnet: " & alumno.n_carnet;
          'Printer.CurrentX = 16.7
          'Printer.Print "PERIODO: " & CONS_NOTA.Combo3.Text
          Printer.Print ""
          Printer.Font.Size = 12
          Printer.Print ""
          Printer.CurrentX = 0.5
          Printer.Print "A R E A S";
          Printer.CurrentX = 6
          Printer.Print "O B S E R V A C I O N E S"
          Printer.Font.Size = 8
          Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
          Printer.CurrentY = 4.9
       End If
    Else
       Printer.Print ""
       'Printer.Print ""
    End If
    End If
    Next REC
    L = L + 8
    Call MODU
    Printer.FontUnderline = True
    Printer.CurrentX = 0.3
    Printer.Print "Indicadores"
    Printer.FontUnderline = False
    Printer.CurrentX = 0.3
    Printer.Print "R=Logro recuperado"
    Printer.CurrentX = 0.3
    Printer.Print "D=Logro no recuperado"
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Line (3.5, Printer.CurrentY)-(8.5, Printer.CurrentY)
    Printer.Line (14, Printer.CurrentY)-(19, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.CurrentX = 4.5
    Printer.Print "Marcela de la Torre";
    Printer.CurrentX = 15.5
    Printer.Print "Mauricio Roa"
    Printer.CurrentX = 5.2
    Printer.Print vini.VRector;
    Printer.CurrentX = 14.8
    Printer.Print "Coordinador académico"
    
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
'Ruta = "c:\windows\datos\"
If (Dir(Ruta & "infcur.edu") <> "") Then
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo1.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Combo1.Text = Combo1.List(0)
End If

Option1.Value = True
Option3.Value = True
Text1.MaxLength = 2
Text2.MaxLength = 2
'Frame1.Caption = CONS_NOTA.Combo4.Text
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
If (L > 60) Then
               Printer.Line (4.8, 4.2)-(4.8, Printer.CurrentY)
               Printer.Line (5.6, 4.2)-(5.6, Printer.CurrentY)
               'Printer.Line (6.4, 4.2)-(6.4, Printer.CurrentY)
               'Printer.Line (7.2, 4.2)-(7.2, Printer.CurrentY)
               'Printer.Line (8, 4.2)-(8, Printer.CurrentY)
               Printer.NewPage
               L = 1
               NAR = FreeFile
               Open Ruta & "inicial.edu" For Input As #NAR
               Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
               Close #NAR
               Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
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
                  Printer.Print Combo1.Text;
                  Printer.CurrentX = 13.5
                  Printer.Print alumno.n_carnet;
                  Printer.CurrentX = 18.3
                  'Printer.Print CONS_NOTA.Combo3.Text
                  Printer.CurrentY = 4.9
               Else
                  Printer.Font.Size = 14
                  Printer.CurrentY = 1
                  Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
                  Printer.FontBold = True
                  Printer.Print ini.nombre
                  Printer.CurrentX = 6.2
                  Printer.Print "REPORTE DE RECUPERACIONES"
                  Printer.FontBold = False
                  Printer.Print ""
                  Printer.Font.Size = 10
                  Printer.CurrentX = 0.5
                  Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & VV & ").";
                  Printer.CurrentX = 16.7
                  Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
                  'Printer.CurrentX = 0.5
                  'Printer.Print "GRADO: " & RE22;
                  'Printer.CurrentX = 6
                  Printer.CurrentX = 0.5
                  Printer.Print "GRUPO: " & Combo1.Text;
                  Printer.CurrentX = 12.7
                  Printer.Print "No.carnet: " & alumno.n_carnet;
                  'Printer.CurrentX = 16.7
                  'Printer.Print "PERIODO: " & CONS_NOTA.Combo3.Text
                  Printer.Print ""
                  Printer.Font.Size = 12
                  Printer.Print ""
                  Printer.CurrentX = 0.5
                  Printer.Print "A R E A S";
                  Printer.CurrentX = 6
                  Printer.Print "O B S E R V A C I O N E S"
                  Printer.Font.Size = 8
                  Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
                  Printer.CurrentY = 4.9
               End If
            Printer.Font.Size = 8
End If
End Sub
