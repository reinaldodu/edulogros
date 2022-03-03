VERSION 5.00
Begin VB.Form FORMATO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "FORMATO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command3 
         Caption         =   "-"
         Height          =   320
         Left            =   2130
         TabIndex        =   7
         Top             =   720
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   320
         Left            =   1935
         TabIndex        =   6
         Top             =   720
         Width           =   195
      End
      Begin VB.TextBox Txt_Espa 
         Height          =   320
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Encabezado"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sin Formato"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Con Formato"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FORMATO"
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
'Dim profe As maestropro
'Dim icur As inforcur
'Dim ini As inicio
'Dim leye As leyendis
RESP = MsgBox("DESEA IMPRIMIR EL REPORTE QUE APARECE EN PANTALLA?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR REPORTE")
If RESP = vbYes Then
    Screen.MousePointer = 11
    L = 0
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
            If RTrim(icur.nom) = CONS_NOTA.Combo4.Text Then
                SP = RTrim(icur.director)
            End If
    Wend
    Close #NAR
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
    Close #NAR
    Printer.ScaleMode = 7
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
       Printer.Print "ESTUDIANTE: " & CONS_NOTA.Frame1.Caption;
       'Printer.CurrentX = 12.7
       'Printer.Print "CIUDAD: " & ini.ciudad;
       Printer.CurrentX = 16.7
       Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
       Printer.CurrentX = 0.5
       Printer.Print "GRADO: " & CONS_NOTA.Label9.Caption;
       Printer.CurrentX = 6
       Printer.Print "GRUPO: " & CONS_NOTA.Label11.Caption;
       Printer.CurrentX = 12.7
       Printer.Print "No.carnet: " & CONS_NOTA.Label17.Caption;
       Printer.CurrentX = 16.7
       Printer.Print "PERIODO: " & CONS_NOTA.Label10.Caption
       Printer.Print ""
       Printer.Font.Size = 12
       Printer.CurrentX = 0.5
       Printer.Print "MATERIAS";
       Printer.CurrentX = 5
       Printer.Print "DESP";
       Printer.CurrentX = 5.7
       Printer.Print "%";
       'Printer.CurrentX = 6.5
       'Printer.Print "J.V";
       'Printer.CurrentX = 7.3
       'Printer.Print "IND";
       'Printer.CurrentX = 8.2
       'Printer.Print "O B S E R V A C I O N E S"
       'Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
       'Printer.CurrentY = 4.9
    Else
       If Check1.Value = 1 Then
          Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.9
       Else
          Printer.CurrentY = 2.9
       End If
       Printer.Font.Size = 10
       Printer.CurrentX = 3.2
       Printer.Print CONS_NOTA.Frame1.Caption;
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
       Printer.Print CONS_NOTA.Label9.Caption;
       Printer.CurrentX = 7.5
       Printer.Print CONS_NOTA.Label11.Caption;
       Printer.CurrentX = 13.5
       Printer.Print CONS_NOTA.Label17.Caption;
       Printer.CurrentX = 18.3
       Printer.Print CONS_NOTA.Label10.Caption
       Printer.CurrentY = 4.9
    End If
    Printer.Font.Size = 8
    k = 1
    While (k < CONS_NOTA.MATI20.Rows - 1)
        If CONS_NOTA.MATI20.TextMatrix(k, 1) <> "" Then
            Printer.CurrentX = 0.3
            Printer.Print CONS_NOTA.MATI20.TextMatrix(k, 0);
            Printer.CurrentX = 5.1
            Printer.Print CONS_NOTA.MATI20.TextMatrix(k, 1);
            Printer.CurrentX = 5.8
            Printer.Print CONS_NOTA.MATI20.TextMatrix(k, 2);
            Printer.CurrentX = 6.6
            Printer.Print CONS_NOTA.MATI20.TextMatrix(k, 3);
            Printer.CurrentX = 7.5
            Printer.Print CONS_NOTA.MATI20.TextMatrix(k, 4);
            X = 92
            L1 = Left(CONS_NOTA.MATI20.TextMatrix(k, 5), X)
            While Right(L1, 1) <> " "
                    If X = 1 Then
                        GoTo tero
                    End If
                    X = X - 1
                    L1 = Left(L1, X)
            Wend
tero:
            Printer.CurrentX = 8.2
            Printer.Print L1
            L = L + 1
            Call MODULA
            Y = Len(L1)
            Y = 200 - Y
            L2 = Right(CONS_NOTA.MATI20.TextMatrix(k, 5), Y)
            If RTrim(L2) <> "" Then
               Printer.CurrentX = 8.2
               Printer.Print L2
               L = L + 1
               Call MODULA
            End If
        Else
            Printer.CurrentX = 7.5
            Printer.Print CONS_NOTA.MATI20.TextMatrix(k, 4);
            Printer.CurrentX = 8.2
            X = 92
            L1 = Left(CONS_NOTA.MATI20.TextMatrix(k, 5), X)
            While Right(L1, 1) <> " "
                If X = 1 Then
                    GoTo tero2
                End If
                X = X - 1
                L1 = Left(L1, X)
            Wend
tero2:
            Printer.CurrentX = 8.2
            Printer.Print L1
            L = L + 1
            Call MODULA
            Y = Len(L1)
            Y = 200 - Y
            L2 = Right(CONS_NOTA.MATI20.TextMatrix(k, 5), Y)
            If RTrim(L2) <> "" Then
               Printer.CurrentX = 8.2
               Printer.Print L2
               L = L + 1
               Call MODULA
            End If
        End If
        k = k + 1
        If CONS_NOTA.MATI20.TextMatrix(k, 3) <> "" Then
            Printer.Print ""
            L = L + 1
            Call MODULA
            Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
            Printer.Print ""
            L = L + 1
            Call MODULA
        End If
    Wend
    Printer.Print ""
    L = L + 1
    Call MODULA
    Printer.Line (0.2, Printer.CurrentY)-(20.2, Printer.CurrentY)
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
          Printer.Print CONS_NOTA.Frame1.Caption;
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
          Printer.Print CONS_NOTA.Label9.Caption;
          Printer.CurrentX = 7.5
          Printer.Print CONS_NOTA.Label11.Caption;
          Printer.CurrentX = 13.5
          Printer.Print CONS_NOTA.Label17.Caption;
          Printer.CurrentX = 18.3
          Printer.Print CONS_NOTA.Label10.Caption
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
          Printer.Print "ESTUDIANTE: " & CONS_NOTA.Frame1.Caption;
          'Printer.CurrentX = 12.7
          'Printer.Print "CIUDAD: " & ini.ciudad;
          Printer.CurrentX = 16.7
          Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
          Printer.CurrentX = 0.5
          Printer.Print "GRADO: " & CONS_NOTA.Label9.Caption;
          Printer.CurrentX = 6
          Printer.Print "GRUPO: " & CONS_NOTA.Label11.Caption;
          Printer.CurrentX = 12.7
          Printer.Print "No.carnet: " & CONS_NOTA.Label17.Caption;
          Printer.CurrentX = 16.7
          Printer.Print "PERIODO: " & CONS_NOTA.Label10.Caption
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
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    Get #NAR, SP, profe
    Close #NAR
    Printer.CurrentX = 15.6 - ((Len(RTrim(profe.nombres) & " " & RTrim(profe.apellidos)) / 4.8) / 2)
    Printer.Print RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
    
    Printer.CurrentX = 4.5
    Printer.Print vini.VRector;
    
    Printer.CurrentX = 14
    Printer.Print vini.VDirector
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
Check1.Value = 0
Txt_Espa.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub MODULA()
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
               If Option1.Value = True Then
                  If Check1.Value = 1 Then
                     Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.9
                  Else
                     Printer.CurrentY = 2.9
                  End If
                  Printer.Font.Size = 10
                  Printer.CurrentX = 3.2
                  Printer.Print CONS_NOTA.Frame1.Caption;
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
                  Printer.Print CONS_NOTA.Label9.Caption;
                  Printer.CurrentX = 7.5
                  Printer.Print CONS_NOTA.Label11.Caption;
                  Printer.CurrentX = 13.5
                  Printer.Print CONS_NOTA.Label17.Caption;
                  Printer.CurrentX = 18.3
                  Printer.Print CONS_NOTA.Label10.Caption
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
                  Printer.Print "ESTUDIANTE: " & CONS_NOTA.Frame1.Caption;
                  'Printer.CurrentX = 12.7
                  'Printer.Print "CIUDAD: " & ini.ciudad;
                  Printer.CurrentX = 16.7
                  Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
                  Printer.CurrentX = 0.5
                  Printer.Print "GRADO: " & CONS_NOTA.Label9.Caption;
                  Printer.CurrentX = 6
                  Printer.Print "GRUPO: " & CONS_NOTA.Label11.Caption;
                  Printer.CurrentX = 12.7
                  Printer.Print "No.carnet: " & CONS_NOTA.Label17.Caption;
                  Printer.CurrentX = 16.7
                  Printer.Print "PERIODO: " & CONS_NOTA.Label10.Caption
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
               End If
               Printer.Font.Size = 8
End If
End Sub
