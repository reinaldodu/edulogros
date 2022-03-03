VERSION 5.00
Begin VB.Form AREAS_GRADO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materias por grupo"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   Icon            =   "AREAS_GRADO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "&ELIMINAR"
      Height          =   735
      Left            =   4320
      Picture         =   "AREAS_GRADO.frx":044A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Elimina el área del grupo seleccionado"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   320
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "AREAS_GRADO.frx":054C
      Left            =   960
      List            =   "AREAS_GRADO.frx":057A
      TabIndex        =   5
      Text            =   "PREJARDIN"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Ok"
      Height          =   320
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&GUARDAR"
      Height          =   735
      Left            =   2760
      Picture         =   "AREAS_GRADO.frx":05FA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guarda el área para el grupo seleccionado"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "ADICIONAR MATERIA"
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5535
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   3975
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PROFESOR"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "INTENSIDAD"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MATERIA"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "IMPRIMIR"
      ForeColor       =   &H00C00000&
      Height          =   1200
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   2415
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GRADO:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Materias adicionadas..."
      Height          =   195
      Left            =   3600
      TabIndex        =   13
      Top             =   2040
      Width           =   1635
   End
End
Attribute VB_Name = "AREAS_GRADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text <> Combo1.List(0) Then
    Combo1.Text = Combo1.List(0)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command4_Click
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo4.SetFocus
End If
End Sub

Private Sub Combo4_Change()
If Combo4.Text <> Combo4.List(0) Then
    Combo4.Text = Combo4.List(0)
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If Command3.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command3_Click
End If
End Sub

Private Sub Command3_Click()
'Dim argra As areagr
'Dim icur As inforcur
'Dim mate As infomater
'Dim profe As maestropro
If Text3.Text = "" Then
MsgBox "ESCRIBA LA INTENSIDAD HORARIA", 16, "MATERIAS POR GRUPO"
    Text3.SetFocus
    Exit Sub
End If
If Dir(Ruta & Combo4.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
If RTrim(icur.nom) = Combo4.Text Then
SAPO = RTrim(icur.grado)
End If
Wend
Close #NAR
Y = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
Y = Y + 1
Get #NAR, Y, mate
If RTrim(mate.nom) = Combo2.Text Then
que = mate.num
End If
Wend
Close #NAR
pio = 0
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
While Not EOF(NAR)
pio = pio + 1
Get #NAR, pio, profe
If (RTrim(profe.nombres) & " " & RTrim(profe.apellidos)) = Combo3.Text Then
p = pio
End If
Wend
Close #NAR
cona = 0
cona2 = 0
Y = 0
NAR = FreeFile
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
cona = cona + 1
Get #NAR, cona, argra
If (argra.num_area = que) And (RTrim(argra.nom_grup) = Combo4.Text) Then
RESP = MsgBox("ESTA MATERIA YA EXISTE PARA ESTE GRUPO, DESEA REMPLAZARLA?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
argra.grado = SAPO
argra.num_area = que
argra.ih = Text3.Text
argra.num_pro = p
argra.nom_grup = Combo4.Text
Put #NAR, cona, argra
End If
Y = 1
End If
If ((RTrim(argra.grado) = SAPO) And (RTrim(argra.nom_grup) = Combo4.Text)) Then
cona2 = cona2 + 1
End If
Wend
Close #NAR
Text7.Text = cona2
If Y = 1 Then
    Combo2.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA GUARDAR LA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
argra.grado = SAPO
argra.num_area = que
argra.ih = Text3.Text
argra.num_pro = p
argra.nom_grup = Combo4.Text
Put #NAR, cona, argra
Close #NAR
Text7.Text = Text7.Text + 1
End If
Combo2.SetFocus
End Sub
Private Sub Command4_Click()
'Dim argra As areagr
'Dim ini As inicio
'Dim mate As infomater
'Dim profe As maestropro
NAR = FreeFile
If Dir(Ruta & "areagra.edu") = "" Then
MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 48, "IMPRIMIR"
Exit Sub
End If
CLIS = 0
OPP = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
CLIS = CLIS + 1
Get #NAR, CLIS, argra
If RTrim(argra.grado) = RTrim(Combo1.Text) Then
OPP = 1
End If
Wend
Close #NAR
If OPP = 0 Then
MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 48, "IMPRIMIR"
Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR LAS MATERIAS DE ESTE GRADO?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
Printer.ScaleMode = 7
PAG = 1
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "Pág." & PAG
Printer.CurrentY = 2
Printer.CurrentX = 8
Printer.Font.Size = 10
Printer.Font.Underline = True
Printer.Print "LISTA DE MATERIAS POR GRUPO"
Printer.CurrentY = 3
Printer.CurrentX = 2.5
Printer.Font.Underline = False
Printer.Print ini.nombre
Printer.CurrentY = 3.5
Printer.CurrentX = 2.5
Printer.Print "GRADO    : " & RTrim(Combo1.Text)
Printer.Print ""
Printer.CurrentX = 2.5
Printer.Font.Underline = True
Printer.Print "#MATERIA";
Printer.CurrentX = 4.5
Printer.Print "NOMBRE DE LA MATERIA";
Printer.CurrentX = 10
Printer.Print "I.H";
Printer.CurrentX = 11
Printer.Print "#PR";
Printer.CurrentX = 12
Printer.Print "NOMBRE DEL PROFESOR";
Printer.CurrentX = 17.5
Printer.Print "GRUPO"
Printer.Font.Size = 8
Printer.Font.Underline = False
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
clu = 0
CHA = 0
While Not EOF(NAR)
clu = clu + 1
Get #NAR, clu, argra
If RTrim(argra.grado) = RTrim(Combo1.Text) Then
Printer.CurrentX = 2.5
Printer.Print argra.num_area;
Printer.CurrentX = 4.5
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
Get #NAR, (argra.num_area), mate
Close #NAR
Printer.Print mate.nom;
Printer.CurrentX = 10
Printer.Print argra.ih;
Printer.CurrentX = 11
Printer.Print argra.num_pro;
Printer.CurrentX = 12
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, (argra.num_pro), profe
Close #NAR
Printer.Print RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
Printer.CurrentX = 17.5
Printer.Print RTrim(argra.nom_grup)
CHA = CHA + 1
NAR = NAR - 1
Else
GoTo jojo
End If
If (CHA Mod 60) = 0 Then
Printer.NewPage
PAG = PAG + 1
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "Pág." & PAG
Printer.CurrentY = 2
Printer.CurrentX = 8
Printer.Font.Size = 10
Printer.Font.Underline = True
Printer.Print "LISTA DE MATERIAS POR GRUPO"
Printer.CurrentY = 3
Printer.CurrentX = 2.5
Printer.Font.Underline = False
Printer.Print ini.nombre
Printer.CurrentY = 3.5
Printer.CurrentX = 2.5
Printer.Print "GRADO    : " & RTrim(Combo1.Text)
Printer.Print ""
Printer.CurrentX = 2.5
Printer.Font.Underline = True
Printer.Print "#MATERIA";
Printer.CurrentX = 4.5
Printer.Print "NOMBRE DE LA MATERIA";
Printer.CurrentX = 10
Printer.Print "I.H";
Printer.CurrentX = 11
Printer.Print "#PR";
Printer.CurrentX = 12
Printer.Print "NOMBRE DEL PROFESOR";
Printer.CurrentX = 17.5
Printer.Print "GRUPO"
Printer.Font.Size = 8
Printer.Font.Underline = False
End If
jojo:
Wend
Close #NAR
Printer.EndDoc
End If
End Sub

Private Sub Command5_Click()
'Dim argra As areagr
'Dim mate As infomater
RESP = MsgBox("DESEA ELIMINAR LA MATERIA " & Combo2.Text & " DEL GRUPO " & Combo4.Text & "?", vbYesNo + vbQuestion + vbDefaultButton2, "ELIMINAR MATERIA")
If RESP = vbYes Then
    NAR = FreeFile
    Y = 0
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    While Not EOF(NAR)
        Y = Y + 1
        Get #NAR, Y, mate
        If RTrim(mate.nom) = Combo2.Text Then
            que = mate.num
        End If
    Wend
    Close #NAR
    CLO = 0
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        CLO = CLO + 1
        Get #NAR, CLO, argra
        If ((argra.num_area = que) And (RTrim(argra.nom_grup) = Combo4.Text)) Then
            argra.grado = ""
            argra.ih = 0
            argra.nom_grup = ""
            argra.num_area = 0
            argra.num_pro = 0
            Put #NAR, CLO, argra
            Close #NAR
            MsgBox "LA MATERIA HA SIDO ELIMINADA", 16, "MATERIAS POR GRUPO"
            Exit Sub
        End If
    Wend
    Close #NAR
    MsgBox "LA MATERIA NO ESTA CREADA PARA ESTE GRUPO", 64, "ADVERTENCIA"
    Combo2.SetFocus
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Creación de áreas por grupo."
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo3.SetFocus
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

Private Sub Form_Load()
'Dim mate As infomater
'Dim icur As inforcur
'Dim profe As maestropro
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "prinpro.edu") <> "") Then
Command3.Enabled = True
Command5.Enabled = True
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
que = 0
While Not EOF(NAR)
que = que + 1
Get #NAR, que, mate
Wend
Close #NAR
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
For I = 1 To que - 1
Get #NAR, I, mate
If RTrim(mate.nom) <> "" Then
Combo2.AddItem RTrim(mate.nom)
End If
Next I
Close #NAR
pio = 0
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
While Not EOF(NAR)
pio = pio + 1
Get #NAR, pio, profe
Wend
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
For J = 1 To pio - 1
Get #NAR, J, profe
If RTrim(profe.nombres) <> "" Then
Combo3.AddItem RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
End If
Next J
Close #NAR
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
Combo4.AddItem RTrim(icur.nom)
Wend
Close #NAR
Combo2.Text = Combo2.List(0)
Combo3.Text = Combo3.List(0)
Combo4.Text = Combo4.List(0)
Else
Command3.Enabled = False
Command5.Enabled = False
End If
Text3.MaxLength = 2
End Sub
