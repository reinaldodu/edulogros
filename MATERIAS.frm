VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MATERIAS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear materias"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   Icon            =   "MATERIAS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "MATERIAS POR GRUPO"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&CORREGIR"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CO&NSULTAR"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "ADICIONAR MATERIA"
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "&ACEPTAR"
         Height          =   320
         Left            =   1080
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRE:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   120
      Picture         =   "MATERIAS.frx":0442
      ScaleHeight     =   4275
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "MATERIAS"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin MSFlexGridLib.MSFlexGrid MATI5 
         Height          =   1215
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
   End
End
Attribute VB_Name = "MATERIAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim mate As infomater
If RTrim(Text1.Text) = "" Then
    MsgBox "ESCRIBA EL NOMBRE DEL AREA", 16, "ADVERTENCIA"
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
que = 0
While Not EOF(NAR)
    que = que + 1
    Get #NAR, que, mate
    If RTrim(Text1.Text) = RTrim(mate.nom) Then
        MsgBox "EL AREA YA EXISTE", 16, "ADVERTENCIA"
        Close #NAR
        Text1.Text = ""
        Exit Sub
    End If
Wend
Close #NAR
MATI5.Rows = MATI5.Rows + 1
MATI5.TextMatrix((MATI5.Rows - 1), 0) = que
MATI5.TextMatrix((MATI5.Rows - 1), 1) = Format(Text1.Text, ">")
mate.nom = RTrim(Format(Text1.Text, ">"))
mate.num = que
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
Put #NAR, que, mate
Close #NAR
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Unload Me
AREAS_GRADO.Show
End Sub

Private Sub Command4_Click()
CONS_MATER.Show
End Sub

Private Sub Command5_Click()
CORR_MATER.Show 1
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Creación de áreas."
End Sub

Private Sub Form_Load()
Text1.MaxLength = 50
MATI5.Row = 0
MATI5.Col = 0
MATI5.ColWidth(0) = 400
MATI5.CellForeColor = RGB(255, 255, 255)
MATI5.CellBackColor = RGB(0, 0, 150)
MATI5.Text = "No."
MATI5.Col = 1
MATI5.ColWidth(1) = 2400
MATI5.CellForeColor = RGB(255, 255, 255)
MATI5.CellBackColor = RGB(0, 0, 150)
MATI5.Text = "NOMBRE"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
