VERSION 5.00
Begin VB.Form Imp_Lisnota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión libro de notas"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "Imp_Lisnota.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3255
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2700
         TabIndex        =   8
         Top             =   645
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   645
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Códigos"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
         Height          =   195
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox List_grupo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "Imp_Lisnota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Dir(Ruta & List_grupo.Text & ".gru") = "" Then
    MsgBox "NO EXISTE ESTE GRUPO", 48, "LIBRO DE NOTAS"
    List_grupo.SetFocus
    Exit Sub
End If
NAR = FreeFile
ret = 0
Open Ruta & List_grupo.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
If Option1.Value = True Then
    s = 1
    q = ret - 1
    MS1 = "DESEA IMPRIMIR LOS INFORMES FINALES DEL GRUPO " & List_grupo.Text & "?"
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
    MS1 = "DESEA IMPRIMIR LOS INFORMES FINALES DEL GRUPO " & List_grupo.Text & ", DESDE EL CODIGO " & Text1.Text & " HASTA EL CODIGO " & Text2.Text & "?"
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "LIBRO DE NOTAS")
If RESP = vbYes Then
    Screen.MousePointer = 11
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
    Close #NAR
    Printer.ScaleMode = 7
    For VV = s To q
        RECO = False
        Open Ruta & List_grupo.Text & ".gru" For Random As #NAR Len = Len(alugru)
        Get #NAR, VV, alugru
        Close #NAR
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        Printer.Font.Size = 14
        Printer.CurrentY = 3
        Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
        Printer.FontBold = True
        Printer.Print ini.nombre
        'Printer.Font.Size = 11
        'Printer.CurrentX = 10.2 - (Len(ini.Rector) / 5.2) / 2
        'Printer.Print ini.Rector
        'Printer.Print ""
        'Printer.Print ""
        Printer.Font.Size = 14
        Printer.CurrentX = 10.2 - ((Len("INFORME FINAL DE CALIFICACIONES") / 3.3) / 2)
        Printer.Print "INFORME FINAL DE CALIFICACIONES"
        Printer.FontBold = False
        Printer.Print ""
        Printer.Print ""
        Printer.Font.Size = 12
        Printer.CurrentX = 3
        Printer.Print "AÑO LECTIVO: " & Year(Date);
        Printer.CurrentX = 16
        Printer.Print "CURSO: " & Right(List_grupo.Text, Len(List_grupo.Text) - 1)
        Printer.CurrentX = 3
        Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
        Printer.CurrentX = 16
        Printer.Print "No.carnet: " & alumno.n_carnet
        Printer.Print ""
        Printer.Print ""
        Printer.FontBold = True
        Printer.CurrentX = 3
        Printer.Print "A R E A S";
        Printer.CurrentX = 16
        Printer.Print "VALORACION FINAL"
        Printer.FontBold = False
        Printer.Print ""
        cona = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            cona = cona + 1
            Get #NAR, cona, argra
            If RTrim(argra.nom_grup) = List_grupo.Text Then
                NAR = FreeFile
                Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                Get #NAR, argra.num_area, mate
                Close #NAR
                NAR = NAR - 1
                Printer.CurrentX = 3
                Printer.Print RTrim(mate.nom);
                If Dir(Ruta & List_grupo.Text & argra.num_area & "5.obs") <> "" Then
                        NAR = FreeFile
                        z = 0
                        Open Ruta & List_grupo.Text & argra.num_area & "5.obs" For Random As #NAR Len = Len(notas)
                        While Not EOF(NAR)
                            z = z + 1
                            Get #NAR, z, notas
                            If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                                Printer.CurrentX = 17
                                Printer.Print notas.FA
                                GoTo camila
                            End If
                        Wend
camila:
                        Close #NAR
                        NAR = NAR - 1
                Else
                    Printer.Print ""
                End If
            End If
        Wend
        Close #NAR
        Printer.CurrentY = 20
        Printer.CurrentX = 3
        Printer.Print "OBSERVACIONES:"
        Printer.Line (7, Printer.CurrentY)-(20, Printer.CurrentY)
        Printer.Line (3, Printer.CurrentY + 0.5)-(20, Printer.CurrentY + 0.5)
        Printer.Line (3, Printer.CurrentY + 0.5)-(20, Printer.CurrentY + 0.5)
        Printer.Line (3, 25)-(8.5, 25)
        Printer.Line (13, 25)-(18.5, 25)
        Printer.CurrentY = 25.2
        Printer.CurrentX = 4.3
        Printer.Print vini.VRector;
        Printer.CurrentX = 13.9
        Printer.Print "Firma del Secretario(a)."
        Printer.NewPage
    Next VV
    Printer.EndDoc
    Printer.Font.Size = 8
    Unload Me
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
If Dir(Ruta & "infcur.edu") <> "" Then
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        List_grupo.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    List_grupo.Text = List_grupo.List(0)
Else
    Command1.Enabled = False
End If
Option1.Value = True
Text1.MaxLength = 2
Text2.MaxLength = 2
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
