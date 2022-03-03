VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GRUPO_GRA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos existentes por grado"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8550
   Icon            =   "GRUPO_GRA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
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
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "GRUPO_GRA.frx":044A
         Left            =   960
         List            =   "GRUPO_GRA.frx":0478
         TabIndex        =   1
         Text            =   "PREJARDIN"
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "GRUPO_GRA.frx":04F8
         Left            =   960
         List            =   "GRUPO_GRA.frx":0508
         TabIndex        =   0
         Text            =   "UNICA"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GRADO:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   810
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupos"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5415
      Begin MSFlexGridLib.MSFlexGrid MATI16 
         Height          =   1335
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total de grupos..."
      Height          =   195
      Left            =   6600
      TabIndex        =   7
      Top             =   2280
      Width           =   1245
   End
End
Attribute VB_Name = "GRUPO_GRA"
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
Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If Command2.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command2_Click
End If
End Sub

Private Sub Command1_Click()
'Dim ini As inicio
If CUCU = 0 Then
    MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 32, "IMPRIMIR"
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
Printer.ScaleMode = 7
Printer.Font.Size = 12
Printer.Font.Underline = True
Printer.CurrentY = 2
Printer.CurrentX = 3
Printer.Print "GRUPOS EXISTENTES DE LA JORNADA DE LA " & Combo1.Text & " GRADO " & Combo2.Text
Printer.CurrentY = 3
Printer.CurrentX = 3
Printer.Font.Underline = False
Printer.Print ini.nombre
Printer.CurrentY = 4
Printer.CurrentX = 3
Printer.Print "GRUPO";
Printer.CurrentX = 7
Printer.Print "DIRECTOR"
Printer.Font.Size = 10
For DF = 1 To CUCU
MATI16.Row = DF
MATI16.Col = 0
Printer.CurrentX = 3
Printer.Print MATI16.Text;
MATI16.Col = 1
Printer.CurrentX = 7
Printer.Print MATI16.Text
Next DF
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Command2_Click()
'Dim icur As inforcur
'Dim profe As maestropro
MATI16.Rows = 1
CUCU = 0
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If (RTrim(icur.jornada) = RTrim(Combo1.Text)) And (RTrim(icur.grado) = RTrim(Combo2.Text)) Then
        NAR = FreeFile
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, (icur.director), profe
        Close #NAR
        CUCU = CUCU + 1
        MATI16.Rows = CUCU + 1
        MATI16.TextMatrix(CUCU, 0) = icur.nom
        MATI16.TextMatrix(CUCU, 1) = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        NAR = NAR - 1
    End If
Wend
Close #NAR
If CUCU = 0 Then
    MsgBox "NO EXISTEN GRUPOS EN ESTA JORNADA Y GRADO", 48, "GRUPOS POR GRADO"
End If
Text1.Text = CUCU
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra los grupos existentes en cada grado, con su respectivo director de grupo."
End Sub

Private Sub Form_Load()
MATI16.Row = 0
MATI16.Col = 0
MATI16.ColWidth(0) = 1500
MATI16.CellFontBold = True
MATI16.CellForeColor = RGB(0, 0, 255)
MATI16.Text = "GRUPO"
MATI16.Col = 1
MATI16.ColWidth(1) = 3000
MATI16.CellFontBold = True
MATI16.CellForeColor = RGB(0, 0, 255)
MATI16.Text = "DIRECTOR"
If Dir(Ruta & "infcur.edu") = "" Then
    Command1.Enabled = False
    Command2.Enabled = False
Else
    Command1.Enabled = True
    Command2.Enabled = True
End If
CUCU = 0
End Sub
