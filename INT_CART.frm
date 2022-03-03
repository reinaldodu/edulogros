VERSION 5.00
Begin VB.Form INT_CART 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intervalo de impresión"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "INT_CART.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "&R.H."
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Códigos"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
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
         TabIndex        =   4
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicial..."
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   540
      End
   End
End
Attribute VB_Name = "INT_CART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim alugru As grupoalu
'Dim ini As inicio
If Option1.Value = True Then
    RESP = MsgBox("DESEA IMPRIMIR TODOS LOS CARNETS DE ESTE GRUPO?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR CARNET")
    If RESP = vbYes Then
        Screen.MousePointer = 11
        Printer.Orientation = 2
        Printer.PaperSize = 5
        NAR = FreeFile
        Open Ruta & "inicial.edu" For Input As #NAR
        Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
        Close #NAR
        Printer.ScaleMode = 7
        CY = 0
        For I = 1 To ret - 1
            CX = 0
            If (I Mod 2) = 0 Then
                CX = 16.4
            End If
            Printer.CurrentY = 0 + CY
            Printer.CurrentX = 0.5 + CX
            Printer.FontBold = True
            Printer.Font.Size = 10
            Printer.Print ini.nombre
            Printer.FontBold = False
            Printer.Font.Size = 8
            Printer.CurrentX = 0.5 + CX
            Printer.Print "Teléfonos: " & ini.Telefono;
            Printer.CurrentX = 8.7 + CX
            Printer.Print "CIUDAD: " & ini.ciudad
            Open Ruta & CARNET.Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
            Get #NAR, I, alugru
            Close #NAR
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, (Val(alugru.num_carnet)), alumno
            Close #NAR
            Printer.CurrentY = 1.5 + CY
            Printer.CurrentX = 3.3 + CX
            Printer.Print "CARNET No." & alumno.n_carnet;
            Printer.CurrentX = 8.7 + CX
            Printer.Print "AÑO ESCOLAR: " & CARNET.Text3.Text
            Printer.Print ""
            Printer.CurrentX = 3.3 + CX
            Printer.Print "NIVEL: " & CARNET.Combo1.Text;
            Printer.CurrentX = 8.7 + CX
            Printer.Print "TELEFONO: " & alumno.tel_acu
            Printer.Print ""
            Printer.CurrentX = 3.3 + CX
            Printer.Print "IDENTIFICACION";
            Printer.CurrentX = 8.7 + CX
            If Check1.Value = 1 Then
                Printer.Print "R-H: " & alumno.rh
            Else
                Printer.Print ""
            End If
            Printer.CurrentX = 3.3 + CX
            Printer.Print "No." & alumno.documento
            Printer.CurrentY = 3.4 + CY
            Printer.CurrentX = 8.7 + CX
            Printer.Print "VALIDO HASTA:"
            Printer.CurrentY = 3.7 + CY
            Printer.CurrentX = 0.5 + CX
            Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
            Printer.CurrentX = 8.7 + CX
            Printer.Print CARNET.Text4.Text;
            Printer.CurrentX = 12.5 + CX
            Printer.Print "FIRMA AUTORIZADA"
            If (I Mod 2) = 0 Then
                CY = CY + 5.4
            End If
            If (I Mod 8) = 0 Then
                Printer.NewPage
                CY = 0
            End If
        Next I
        Printer.EndDoc
        Unload Me
        Printer.Orientation = 1
        Printer.PaperSize = 1
    End If
End If
If Option2.Value = True Then
If Text1.Text = "" Then
    MsgBox "ESCRIBA EL CODIGO INICIAL", 48, "ADVERTENCIA"
    Text1.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "ESCRIBA EL CODIGO FINAL", 48, "ADVERTENCIA"
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
If (Val(Text1.Text) < 1) Or (Val(Text1.Text) >= ret) Then
    MsgBox "NO EXISTE EL CODIGO INICIAL", 48, "ADVERTENCIA"
    Text1.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
If (Val(Text2.Text) < 1) Or (Val(Text2.Text) >= ret) Then
    MsgBox "NO EXISTE EL CODIGO FINAL", 48, "ADVERTENCIA"
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
If Val(Text1.Text) > Val(Text2.Text) Then
    MsgBox "EL CODIGO INICIAL DEBE SER MENOR O IGUAL QUE EL FINAL", 64, "ADVERTENCIA"
    Text1.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR LOS CARNETS DEL CODIGO " & Text1.Text & " AL CODIGO " & Text2.Text & " DE ESTE GRUPO?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR CARNET")
If RESP = vbYes Then
Screen.MousePointer = 11
Printer.Orientation = 2
Printer.PaperSize = 5
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
Printer.ScaleMode = 7
Printer.Font.Size = 8
CY = 0
If (Val(Text1.Text) Mod 2) = 0 Then
h = 1
Else
h = 0
End If
J = 0
For I = Val(Text1.Text) To Val(Text2.Text)
CX = 0
If (I Mod 2) = h Then
CX = 16.4
End If
Printer.CurrentY = 0 + CY
Printer.CurrentX = 0.5 + CX
Printer.FontBold = True
Printer.Font.Size = 10
Printer.Print ini.nombre
Printer.FontBold = False
Printer.Font.Size = 8
Printer.CurrentX = 0.5 + CX
Printer.Print "Teléfonos: " & ini.Telefono
Printer.CurrentX = 8.7 + CX
Printer.Print "CIUDAD: " & ini.ciudad
Open Ruta & CARNET.Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
Get #NAR, I, alugru
Close #NAR
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, (Val(alugru.num_carnet)), alumno
Close #NAR
Printer.CurrentY = 1.5 + CY
Printer.CurrentX = 3.3 + CX
Printer.Print "CARNET No." & alumno.n_carnet;
Printer.CurrentX = 8.7 + CX
Printer.Print "AÑO ESCOLAR: " & CARNET.Text3.Text
Printer.Print ""
Printer.CurrentX = 3.3 + CX
Printer.Print "NIVEL: " & CARNET.Combo1.Text;
Printer.CurrentX = 8.7 + CX
Printer.Print "TELEFONO: " & alumno.tel_acu
Printer.Print ""
Printer.CurrentX = 3.3 + CX
Printer.Print "IDENTIFICACION";
Printer.CurrentX = 8.7 + CX
If Check1.Value = 1 Then
    Printer.Print "R-H: " & alumno.rh
Else
    Printer.Print ""
End If
Printer.CurrentX = 3.3 + CX
Printer.Print "No." & alumno.documento
Printer.CurrentY = 3.4 + CY
Printer.CurrentX = 8.7 + CX
Printer.Print "VALIDO HASTA:"
Printer.CurrentY = 3.7 + CY
Printer.CurrentX = 0.5 + CX
Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
Printer.CurrentX = 8.7 + CX
Printer.Print CARNET.Text4.Text;
Printer.CurrentX = 12.5 + CX
Printer.Print "FIRMA AUTORIZADA"
J = J + 1
If (I Mod 2) = h Then
CY = CY + 5.4
End If
If (J Mod 8) = 0 Then
Printer.NewPage
CY = 0
End If
Next I
Printer.EndDoc
Unload Me
Printer.Orientation = 1
Printer.PaperSize = 1
End If
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Text1.MaxLength = 2
Text2.MaxLength = 2
Option1.Value = True
Check1.Value = 1
Frame1.Caption = "CARNETS GRUPO " & CARNET.Combo1.Text
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
