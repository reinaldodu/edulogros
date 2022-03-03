VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PENGRU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alumnos pendientes y sin grupo"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "PENGRU.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Copiar"
      Height          =   615
      Left            =   4560
      Picture         =   "PENGRU.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Copiar la información que se muestra en pantalla"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   5760
      Picture         =   "PENGRU.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir la información que se muestra en pantalla"
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Or&denar"
      Height          =   615
      Left            =   3360
      Picture         =   "PENGRU.frx":0EA6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ordenar ascendentemente por apellidos"
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "PENGRU.frx":0FA8
         Left            =   120
         List            =   "PENGRU.frx":0FB2
         TabIndex        =   0
         Text            =   "PENDIENTE"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00C00000&
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid MATI21 
         Height          =   3975
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         BackColorBkg    =   12632256
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Estudiantes..."
         Height          =   195
         Left            =   6480
         TabIndex        =   9
         Top             =   4320
         Width           =   1365
      End
   End
End
Attribute VB_Name = "PENGRU"
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
Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim aluper As pertgrup
Screen.MousePointer = 11
MATI21.Rows = 1
J = 1
MS1 = "ESTUDIANTES PENDIENTES DE GRUPO"
If Combo1.Text = "SIN GRUPO" Then
    MS1 = "ESTUDIANTES SIN GRUPO"
End If
Frame1.Caption = MS1
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
For h = 1 To (I - 1)
Get #NAR, h, aluper
If RTrim(aluper.grupo) = RTrim(Combo1.Text) Then
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, h, alumno
    If (RTrim(alumno.nombres) <> "") And (RTrim(alumno.apellidos) <> "") Then
        MATI21.Rows = J + 1
        MATI21.TextMatrix(J, 0) = alumno.n_carnet
        MATI21.TextMatrix(J, 1) = RTrim(alumno.apellidos)
        MATI21.TextMatrix(J, 2) = RTrim(alumno.nombres)
        MATI21.TextMatrix(J, 3) = RTrim(alumno.grado)
        J = J + 1
    End If
    Close #NAR
    NAR = NAR - 1
End If
Next h
Close #NAR
Text1.Text = J - 1
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
If Val(Text1.Text) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA ORDENAR", 32, "ORDENAR"
    Exit Sub
End If
MATI21.Col = 1
MATI21.Sort = 5
End Sub

Private Sub Command3_Click()
'Dim ini As inicio
If Val(Text1.Text) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA IMPRIMIR", 32, "IMPRIMIR"
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
Printer.CurrentX = 7
Printer.Font.Size = 12
Printer.Print Frame1.Caption
Printer.Font.Size = 10
Printer.Print ""
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
Printer.Print "NOMBRES";
Printer.CurrentX = 13
Printer.Print "GRADO"
Printer.Print ""
For z = 1 To Val(Text1.Text)
    Printer.CurrentX = 1
    Printer.Print MATI21.TextMatrix(z, 0);
    Printer.CurrentX = 3
    Printer.Print MATI21.TextMatrix(z, 1);
    Printer.CurrentX = 8
    Printer.Print MATI21.TextMatrix(z, 2);
    Printer.CurrentX = 13
    Printer.Print MATI21.TextMatrix(z, 3)
    If (z Mod 52) = 0 Then
        PAG = PAG + 1
        Printer.NewPage
        Printer.CurrentY = 1
        Printer.CurrentX = 7
        Printer.Font.Size = 12
        Printer.Print Frame1.Caption
        Printer.Font.Size = 10
        Printer.Print ""
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
        Printer.Print "NOMBRES";
        Printer.CurrentX = 13
        Printer.Print "GRADO"
        Printer.Print ""
    End If
Next z
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "TOTAL ESTUDIANTES..." & Text1.Text
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
For X = 1 To (MATI21.Rows - 1)
        ape = RTrim(MATI21.TextMatrix(X, 0))
        nom = RTrim(MATI21.TextMatrix(X, 1)) & " " & RTrim(MATI21.TextMatrix(X, 2))
        If X < 10 Then
           cop = cop + LTrim(Str(X) & "   - " & ape & "  " & nom) & vbCrLf
        Else
           cop = cop + LTrim(Str(X) & " - " & ape & "  " & nom) & vbCrLf
        End If
Next X
Clipboard.SetText cop
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Alumnos pendientes de grupo y los alumnos que se han retirado de los grupos 'Sin Grupo'."
End Sub

Private Sub Form_Load()
MATI21.Row = 0
MATI21.Col = 0
MATI21.CellForeColor = RGB(255, 255, 255)
MATI21.CellBackColor = RGB(0, 0, 150)
MATI21.ColWidth(0) = 1000
MATI21.Text = "CARNET"
MATI21.Col = 1
MATI21.ColWidth(1) = 2700
MATI21.CellForeColor = RGB(255, 255, 255)
MATI21.CellBackColor = RGB(0, 0, 150)
MATI21.Text = "APELLIDOS"
MATI21.Col = 2
MATI21.ColWidth(2) = 2700
MATI21.CellForeColor = RGB(255, 255, 255)
MATI21.CellBackColor = RGB(0, 0, 150)
MATI21.Text = "NOMBRES"
MATI21.Col = 3
MATI21.ColWidth(3) = 1500
MATI21.CellForeColor = RGB(255, 255, 255)
MATI21.CellBackColor = RGB(0, 0, 150)
MATI21.Text = "GRADO"
End Sub
