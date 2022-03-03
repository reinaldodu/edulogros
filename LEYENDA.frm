VERSION 5.00
Begin VB.Form LEYENDA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comentarios en el boletín"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "LEYENDA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Activar las cajas de texto para modificar la información"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      ToolTipText     =   "Guardar comentarios"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   8895
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   8895
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   8895
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   8895
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   8895
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   8895
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   8895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8895
      End
   End
End
Attribute VB_Name = "LEYENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim leye As leyendis
RESP = MsgBox("DESEA GUARDAR LOS CAMBIOS EFECTUADOS?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
    leye.ly1 = RTrim(Text1.Text)
    leye.ly2 = RTrim(Text2.Text)
    leye.ly3 = RTrim(Text3.Text)
    leye.ly4 = RTrim(Text4.Text)
    leye.ly5 = RTrim(Text5.Text)
    leye.ly6 = RTrim(Text6.Text)
    leye.ly7 = RTrim(Text7.Text)
    leye.ly8 = RTrim(Text8.Text)
    NAR = FreeFile
    Open Ruta & "leyenda.edu" For Output As #NAR
    Write #NAR, leye.ly1, leye.ly2, leye.ly3, leye.ly4, leye.ly5, leye.ly6, leye.ly7, leye.ly8
    Close #NAR
End If
ABC = False
End Sub

Private Sub Command2_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Text8.Enabled = True
    Command1.Enabled = True
End If
End Sub

Private Sub Command3_Click()
If ABC = True Then
   Call Command1_Click
   Unload Me
Else
  Unload Me
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Comentarios en el boletín: Esta información aparecerá al final de cada boletín académico."
End Sub

Private Sub Form_Load()
Text1.MaxLength = 130
Text2.MaxLength = 130
Text3.MaxLength = 130
Text4.MaxLength = 130
Text5.MaxLength = 130
Text6.MaxLength = 130
Text7.MaxLength = 130
Text8.MaxLength = 130
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Command1.Enabled = False
ABC = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ABC = True Then
   Call Command1_Click
   Unload Me
Else
  Unload Me
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
ABC = True
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
ABC = True
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
ABC = True
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
ABC = True
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
ABC = True
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
ABC = True
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
ABC = True
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
ABC = True
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text8.SetFocus
End If
End Sub
