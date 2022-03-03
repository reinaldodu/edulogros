VERSION 5.00
Begin VB.Form RECUP_LOGRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECUPERAR"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "RECUP_LOGRO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   320
         Left            =   1920
         TabIndex        =   6
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   320
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.Line Line3 
         DrawMode        =   16  'Merge Pen
         X1              =   120
         X2              =   2520
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código del logro:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Line Line2 
         DrawMode        =   16  'Merge Pen
         X1              =   120
         X2              =   2520
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Número de la observación (1 a 10):"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Line Line1 
         DrawMode        =   16  'Merge Pen
         X1              =   120
         X2              =   2520
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del alumno:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1350
      End
   End
End
Attribute VB_Name = "RECUP_LOGRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ESCRIBA EL CODIGO DEL ALUMNO", 16, "RECUPERAR"
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DE LA OBSERVACION A CORREGIR", 16, "RECUPERAR"
Text2.SetFocus
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "ESCRIBA EL CODIGO DEL LOGRO", 16, "RECUPERAR"
Text3.SetFocus
Exit Sub
End If
If ((Text1.Text > Val(LOGRO_PEN.Label4.Caption)) Or (Text1.Text < 1)) Then
MsgBox "CODIGO DE ALUMNO NO EXISTE", 32, "RECUPERAR"
Text1.SetFocus
Exit Sub
End If
If ((Text2.Text > 10) Or (Text2.Text < 1)) Then
MsgBox "NUMERO DE OBSERVACION NO EXISTE", 48, "RECUPERAR"
Text2.SetFocus
Exit Sub
End If
If Text3.Text >= FERT Then
MsgBox "CODIGO DE LOGRO NO EXISTE", 48, "ADVERTENCIA"
Text3.SetFocus
Exit Sub
End If
LOGRO_PEN.MATI50.Row = Text1.Text
LOGRO_PEN.MATI50.Col = Text2.Text
If LOGRO_PEN.MATI50.Text = "" Then
MsgBox "NO EXISTE CODIGO PARA LA OBSERVACION No." & Text2.Text, 64, "ADVERTENCIA"
Text2.SetFocus
Exit Sub
End If
LOGRO_PEN.MATI50.CellFontBold = True
LOGRO_PEN.MATI50.CellForeColor = RGB(0, 0, 255)
LOGRO_PEN.MATI50.Text = Text3.Text
Text1.SetFocus
End Sub

Private Sub Form_Load()
Text1.MaxLength = 2
Text2.MaxLength = 2
Text3.MaxLength = 2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONCCC25
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONCCC25:
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONCOO25
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONCOO25:
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CCONC25
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CCONC25:
End Sub
