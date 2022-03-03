VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GRADOS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estudiantes existentes por grado"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9390
   Icon            =   "GRADOS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   8160
      Picture         =   "GRADOS.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Copiar la información que se muestra en pantalla"
      Top             =   4200
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   120
      Picture         =   "GRADOS.frx":0974
      ScaleHeight     =   4515
      ScaleWidth      =   1635
      TabIndex        =   13
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      Picture         =   "GRADOS.frx":38EA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir la información que se muestra en pantalla"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      Picture         =   "GRADOS.frx":3E1C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ordenar ascendentemente por apellidos"
      Top             =   4200
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONSULTAR GRADO:"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   1920
      TabIndex        =   10
      Top             =   3960
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "GRADOS.frx":3F1E
         Left            =   3240
         List            =   "GRADOS.frx":3F52
         TabIndex        =   1
         Text            =   "PREJARDIN"
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "GRADOS.frx":3FE8
         Left            =   960
         List            =   "GRADOS.frx":3FF8
         TabIndex        =   0
         Text            =   "UNICA"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "GRADO:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3735
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   7335
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   3240
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid MATI2 
         Height          =   2775
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL ESTUDIANTES..."
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   1845
      End
   End
End
Attribute VB_Name = "GRADOS"
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
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
'Dim alumno As maestroalum
Screen.MousePointer = 11
MATI2.Rows = 1
h = 1
J = 1
Frame1.Caption = "ESTUDIANTES JORNADA " & Combo1.Text & "  GRADO " & Combo2.Text
NAR = FreeFile
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
While Not EOF(NAR)
Get #NAR, h, alumno
If ((RTrim(alumno.jornada) = RTrim(Combo1.Text)) And (RTrim(alumno.grado) = RTrim(Combo2.Text))) Then
MATI2.Rows = J + 1
MATI2.TextMatrix(J, 0) = alumno.n_carnet
MATI2.TextMatrix(J, 1) = RTrim(alumno.apellidos)
MATI2.TextMatrix(J, 2) = RTrim(alumno.nombres)
J = J + 1
End If
h = h + 1
Wend
Close #NAR
Text1.Text = J - 1
Screen.MousePointer = 0
Combo2.SetFocus
End Sub

Private Sub Command2_Click()
If Val(Text1.Text) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA ORDENAR", 32, "ORDENAR"
    Exit Sub
End If
MATI2.Col = 1
MATI2.Sort = 5
End Sub

Private Sub Command3_Click()
'Dim ini As inicio
If Val(Text1.Text) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA IMPRIMIR", 32, "IMPRESION"
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR TODA LA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
PAG = 1
Printer.ScaleMode = 7
Printer.CurrentY = 1
Printer.CurrentX = 6
Printer.Font.Size = 12
Printer.Print Frame1.Caption
Printer.Print ""
Printer.Font.Size = 10
Printer.CurrentX = 1
Printer.Print ini.nombre;
Printer.CurrentX = 19
Printer.Print "Pág." & PAG
Printer.Print ""
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "CARNET";
Printer.CurrentX = 3
Printer.Print "APELLIDOS";
Printer.CurrentX = 8
Printer.Print "NOMBRES"
Printer.Print ""
For z = 1 To Val(Text1.Text)
    Printer.CurrentX = 1
    Printer.Print MATI2.TextMatrix(z, 0);
    Printer.CurrentX = 3
    Printer.Print MATI2.TextMatrix(z, 1);
    Printer.CurrentX = 8
    Printer.Print MATI2.TextMatrix(z, 2)
    If (z Mod 52) = 0 Then
        PAG = PAG + 1
        Printer.NewPage
        Printer.CurrentY = 1
        Printer.CurrentX = 6
        Printer.Font.Size = 12
        Printer.Print Frame1.Caption
        Printer.Print ""
        Printer.Font.Size = 10
        Printer.CurrentX = 1
        Printer.Print ini.nombre;
        Printer.CurrentX = 19
        Printer.Print "Pág." & PAG
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = 1
        Printer.Print "CARNET";
        Printer.CurrentX = 3
        Printer.Print "APELLIDOS";
        Printer.CurrentX = 8
        Printer.Print "NOMBRES"
        Printer.Print ""
    End If
Next z
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "TOTAL ESTUDIANTES..." & Text1.Text;
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Command4_Click()
If Val(Text1.Text) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA COPIAR", 32, "COPIAR"
    Exit Sub
End If
Clipboard.Clear
cop = ""
For X = 1 To (MATI2.Rows - 1)
        ape = RTrim(MATI2.TextMatrix(X, 0))
        nom = RTrim(MATI2.TextMatrix(X, 1)) & " " & RTrim(MATI2.TextMatrix(X, 2))
        If X < 10 Then
           cop = cop + LTrim(Str(X) & "   - " & ape & "  " & nom) & vbCrLf
        Else
           cop = cop + LTrim(Str(X) & " - " & ape & "  " & nom) & vbCrLf
        End If
Next X
Clipboard.SetText cop
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra los alumnos existentes por jornada y grado."
End Sub

Private Sub Form_Load()
MATI2.Row = 0
MATI2.Col = 0
MATI2.CellForeColor = RGB(255, 255, 255)
MATI2.CellBackColor = RGB(0, 0, 150)
MATI2.ColWidth(0) = 1000
MATI2.Text = "CARNET"
MATI2.Col = 1
MATI2.ColWidth(1) = 2700
MATI2.CellForeColor = RGB(255, 255, 255)
MATI2.CellBackColor = RGB(0, 0, 150)
MATI2.Text = "APELLIDOS"
MATI2.Col = 2
MATI2.ColWidth(2) = 2700
MATI2.CellForeColor = RGB(255, 255, 255)
MATI2.CellBackColor = RGB(0, 0, 150)
MATI2.Text = "NOMBRES"
End Sub
