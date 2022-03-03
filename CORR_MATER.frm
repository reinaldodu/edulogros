VERSION 5.00
Begin VB.Form CORR_MATER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Corregir materia"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4830
   Icon            =   "CORR_MATER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
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
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   320
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   3495
      End
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
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Materia No."
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   825
      End
   End
End
Attribute VB_Name = "CORR_MATER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
'Dim mate As infomater
If BOL = False Then
    MsgBox "ESCRIBA EL NUMERO DEL AREA Y DE CLICK EN OK", 32, "GUARDAR MATERIA"
    Text1.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA GUARDAR ESTA MATERIA?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR MATERIA")
If RESP = vbYes Then
    mate.nom = RTrim(Format(Text2.Text, ">"))
    mate.num = cli
    NAR = FreeFile
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    Put #NAR, cli, mate
    Close #NAR
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Luego de haber corregido el nombre del área de click en Guardar."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
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
Private Sub Command1_Click()
'Dim mate As infomater
BOL = False
If Text1.Text = "" Then
    MsgBox "ESCRIBA UN NUMERO DE AREA", 48, "CORREGIR AREA"
    Text2.Text = ""
    Text1.SetFocus
    Exit Sub
End If
cli = Val(Text1.Text)
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
que = 0
While Not EOF(NAR)
    que = que + 1
    Get #NAR, que, mate
Wend
Close #NAR
If ((cli > (que - 1)) Or (cli < 1)) Then
    MsgBox "NO EXISTE EL AREA", 16, "CORREGIR AREA"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Exit Sub
End If
BOL = True
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
Get #NAR, cli, mate
Close #NAR
Text2.Text = RTrim(mate.nom)
Text2.SetFocus
End Sub

Private Sub Form_Load()
Text1.MaxLength = 2
Text2.MaxLength = 50
BOL = False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command2_Click
End If
End Sub
