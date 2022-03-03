VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DIRECT_TEL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directorio telefónico"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "DIRECT_TEL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   320
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   320
      Left            =   3360
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   4935
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   600
      Left            =   5520
      Picture         =   "DIRECT_TEL.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir el directorio telefónico del grupo que aparece en pantalla"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid MATI53 
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   2
         BackColorBkg    =   12632256
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL ESTUDIANTES..."
      Height          =   195
      Left            =   6480
      TabIndex        =   8
      Top             =   480
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DIRECTORIO POR GRUPOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "DIRECT_TEL.frx":0974
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "DIRECT_TEL"
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
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim icur As inforcur
'Dim alugru As grupoalu
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
MATI53.Rows = 1
Frame1.Caption = ""
Text2.Text = 0
leo = 0
NAR = FreeFile
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    leo = leo + 1
    Get #NAR, leo, alugru
Wend
Close #NAR
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
NAR = FreeFile
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
For TN = 1 To leo - 1
    Get #(NAR - 1), TN, alugru
    Get #NAR, (Val(alugru.num_carnet)), alumno
    MATI53.Rows = TN + 1
    MATI53.TextMatrix(TN, 0) = alumno.n_carnet
    MATI53.TextMatrix(TN, 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    MATI53.TextMatrix(TN, 2) = RTrim(alumno.direccion)
    MATI53.TextMatrix(TN, 3) = RTrim(alumno.tel_acu)
    MATI53.TextMatrix(TN, 4) = RTrim(alumno.madre)
    MATI53.TextMatrix(TN, 5) = RTrim(alumno.tel_ma)
    MATI53.TextMatrix(TN, 6) = RTrim(alumno.padre)
    MATI53.TextMatrix(TN, 7) = RTrim(alumno.tel_pa)
Next TN
Close #NAR
Close #(NAR - 1)
Frame1.Caption = "GRUPO " & Combo1.Text
Text2.Text = TN - 1
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
'Dim ini As inicio
If Val(Text2.Text) = 0 Then
    MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 32, "IMPRESION"
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
    Close #NAR
    Printer.ScaleMode = 7
    Printer.Orientation = 2
    Printer.CurrentY = 0
    Printer.CurrentX = 1
    Printer.Font.Size = 10
    Printer.Print "DIRECTORIO TELEFONICO - " & Frame1.Caption
    Printer.CurrentX = 1
    Printer.Print ini.nombre
    Printer.Font.Size = 8
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print "CARNET";
    Printer.CurrentX = 2.5
    Printer.Print "APELLIDOS Y NOMBRES";
    Printer.CurrentX = 7.5
    Printer.Print "DIRECCION";
    Printer.CurrentX = 12.4
    Printer.Print "TELEFONO";
    Printer.CurrentX = 14.1
    Printer.Print "MADRE";
    Printer.CurrentX = 18.8
    Printer.Print "TELEFONO";
    Printer.CurrentX = 20.5
    Printer.Print "PADRE";
    Printer.CurrentX = 25.2
    Printer.Print "TELEFONO"
    Printer.Print ""
    For z = 1 To Val(Text2.Text)
        Printer.Font.Size = 7
        Printer.CurrentX = 1
        Printer.Print MATI53.TextMatrix(z, 0);
        Printer.CurrentX = 2.5
        Printer.Print MATI53.TextMatrix(z, 1);
        Printer.CurrentX = 7.5
        Printer.Print MATI53.TextMatrix(z, 2);
        Printer.CurrentX = 12.4
        Printer.Print MATI53.TextMatrix(z, 3);
        Printer.CurrentX = 14.1
        Printer.Print MATI53.TextMatrix(z, 4);
        Printer.CurrentX = 18.8
        Printer.Print MATI53.TextMatrix(z, 5);
        Printer.CurrentX = 20.5
        Printer.Print MATI53.TextMatrix(z, 6);
        Printer.CurrentX = 25.2
        Printer.Print MATI53.TextMatrix(z, 7)
    Next z
    Printer.Print ""
    Printer.Font.Size = 8
    Printer.CurrentX = 1
    Printer.Print "TOTAL ESTUDIANTES..." & Text2.Text
    Printer.EndDoc
    Printer.Font.Size = 8
    Printer.Orientation = 1
    Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra el directorio telefónico de un grupo determinado."
End Sub

Private Sub Form_Load()
'Dim icur As inforcur
MATI53.Row = 0
MATI53.Col = 0
MATI53.ColWidth(0) = 900
MATI53.CellForeColor = RGB(255, 255, 255)
MATI53.CellBackColor = RGB(0, 0, 150)
MATI53.Text = "CARNET"
MATI53.Col = 1
MATI53.ColWidth(1) = 3700
MATI53.CellForeColor = RGB(255, 255, 255)
MATI53.CellBackColor = RGB(0, 0, 150)
MATI53.Text = "APELLIDOS Y NOMBRES"
MATI53.Col = 2
MATI53.ColWidth(2) = 3700
MATI53.CellForeColor = RGB(255, 255, 255)
MATI53.CellBackColor = RGB(0, 0, 150)
MATI53.Text = "DIRECCION"
MATI53.Col = 3
MATI53.ColWidth(3) = 1600
MATI53.CellForeColor = RGB(255, 255, 255)
MATI53.CellBackColor = RGB(0, 0, 150)
MATI53.Text = "TELEFONO-CASA"
MATI53.Col = 4
MATI53.ColWidth(4) = 3300
MATI53.CellForeColor = RGB(255, 255, 255)
MATI53.CellBackColor = RGB(0, 0, 150)
MATI53.Text = "MADRE"
MATI53.Col = 5
MATI53.ColWidth(5) = 1700
MATI53.CellForeColor = RGB(255, 255, 255)
MATI53.CellBackColor = RGB(0, 0, 150)
MATI53.Text = "TELEFONO-MADRE"
MATI53.Col = 6
MATI53.ColWidth(6) = 3300
MATI53.CellForeColor = RGB(255, 255, 255)
MATI53.CellBackColor = RGB(0, 0, 150)
MATI53.Text = "PADRE"
MATI53.Col = 7
MATI53.ColWidth(7) = 1700
MATI53.CellForeColor = RGB(255, 255, 255)
MATI53.CellBackColor = RGB(0, 0, 150)
MATI53.Text = "TELEFONO-PADRE"
If Dir(Ruta & "infcur.edu") <> "" Then
Command1.Enabled = True
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
End If
End Sub
