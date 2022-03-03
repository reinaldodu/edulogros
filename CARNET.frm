VERSION 5.00
Begin VB.Form CARNET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carnet estudiantíl"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4215
   Icon            =   "CARNET.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Impresión &Individual"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Impresión &Grupal"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3975
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "VALIDO HASTA:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "AÑO ESCOLAR:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO DEL ALUMNO:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRE DEL GRUPO:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1740
      End
   End
End
Attribute VB_Name = "CARNET"
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
Text2.SetFocus
End If
End Sub

Private Sub Command1_Click()
'Dim alugru As grupoalu
If Text3.Text = "" Then
    MsgBox "ESCRIBA EL AÑO ESCOLAR", 64, "CARNET ESTUDIANTIL"
    Text3.SetFocus
    Exit Sub
End If
If Text4.Text = "" Then
    MsgBox "NO HA ESCRITO LA FECHA DE VALIDEZ", 64, "CARNET ESTUDIANTIL"
    Text4.SetFocus
    Exit Sub
End If
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
NAR = FreeFile
ret = 0
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
INT_CART.Show 1
End Sub
Private Sub Command2_Click()
'Dim alumno As maestroalum
'Dim alugru As grupoalu
'Dim ini As inicio
If Text2.Text = "" Then
MsgBox "ESCRIBA EL CODIGO DEL ESTUDIANTE", 64, "CARNET ESTUDIANTIL"
Text2.SetFocus
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "ESCRIBA EL AÑO ESCOLAR", 64, "CARNET ESTUDIANTIL"
Text3.SetFocus
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "NO HA ESCRITO LA FECHA DE VALIDEZ", 64, "CARNET ESTUDIANTIL"
Text4.SetFocus
Exit Sub
End If
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
MsgBox "GRUPO INCORRECTO", 48
Exit Sub
End If
ret = 0
NAR = FreeFile
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
ret = ret + 1
Get #NAR, ret, alugru
Wend
Close #NAR
If (Text2.Text > ret - 1) Or (Text2.Text < 1) Then
MsgBox "CODIGO NO EXISTE EN ESTE GRUPO", 48, "CARNET ESTUDIANTIL"
Text2.SetFocus
Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR EL CARNET DEL ESTUDIANTE?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR CARNET")
If RESP = vbYes Then
Screen.MousePointer = 11
Printer.Orientation = 2
Printer.PaperSize = 5
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
Close #NAR
Printer.ScaleMode = 7
I = Text2.Text
Printer.CurrentY = 0
Printer.CurrentX = 0.5
Printer.FontBold = True
Printer.Font.Size = 10
Printer.Print ini.nombre
Printer.FontBold = False
Printer.Font.Size = 8
Printer.CurrentX = 0.5
Printer.Print "Teléfonos: " & ini.Telefono;
Printer.CurrentX = 8.7
Printer.Print "CIUDAD: " & ini.ciudad
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
Get #NAR, I, alugru
Close #NAR
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, (Val(alugru.num_carnet)), alumno
Close #NAR
Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
Get #NAR, (Val(alugru.num_carnet)), AdiCampo
Close #NAR
Printer.CurrentY = 1.5
Printer.CurrentX = 3.3
Printer.Print "CARNET No." & alumno.n_carnet;
Printer.CurrentX = 8.7
Printer.Print "AÑO ESCOLAR: " & Text3.Text
Printer.Print ""
Printer.CurrentX = 3.3
Printer.Print "NIVEL: " & Combo1.Text;
Printer.CurrentX = 8.7
Printer.Print "TELEFONO: " & AdiCampo.Tel_casa
Printer.Print ""
Printer.CurrentX = 3.3
Printer.Print "IDENTIFICACION";
Printer.CurrentX = 8.7
Printer.Print "R-H: " & alumno.rh
Printer.CurrentX = 3.3
Printer.Print "No." & alumno.documento
Printer.CurrentY = 3.4
Printer.CurrentX = 8.7
Printer.Print "VALIDO HASTA:"
Printer.CurrentY = 3.7
Printer.CurrentX = 0.5
Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
Printer.CurrentX = 8.7
Printer.Print Text4.Text;
Printer.CurrentX = 12.5
Printer.Print "FIRMA AUTORIZADA"
Printer.EndDoc
Printer.Orientation = 1
Printer.PaperSize = 1
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Imprime los carnets de un grupo determinado o el carnet de un solo alumno."
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
End If
End Sub

Private Sub Form_Load()
'Dim icur As inforcur
If Dir(Ruta & "infcur.edu") <> "" Then
Command1.Enabled = True
Command2.Enabled = True
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
Combo1.AddItem RTrim(icur.nom)
Wend
Close #NAR
Combo1.Text = Combo1.List(0)
Else
Command1.Enabled = False
Command2.Enabled = False
End If
Text2.MaxLength = 2
Text3.MaxLength = 15
Text4.MaxLength = 25
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

