VERSION 5.00
Begin VB.Form CORR_OBSER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CORREGIR OBSERVACION"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CORR_OBSER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
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
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   375
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
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   375
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
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ACEPTAR"
      Height          =   375
      Left            =   1080
      Picture         =   "CORR_OBSER.frx":0442
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   2880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   2880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   2880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   2880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "CORR_OBSER.frx":0884
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "NUEVO CODIGO DE OBSERVACION..."
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CODIGO DEL ALUMNO:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "NUMERO DE LA OBSERVACION (1 A 10):"
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2145
   End
End
Attribute VB_Name = "CORR_OBSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ESCRIBA EL CODIGO DEL ALUMNO", 16, "CORREGIR"
Text1.SetFocus
GoTo NOV
End If
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DE LA OBSERVACION A CORREGIR", 16, "CORREGIR"
Text2.SetFocus
GoTo NOV
End If
If Text3.Text = "" Then
MsgBox "ESCRIBA EL NUEVO CODIGO DE LA OBSERVACION", 16, "CORREGIR"
Text3.SetFocus
GoTo NOV
End If
If ((Text1.Text > GRABAR_OBSER.Text7.Text) Or (Text1.Text < 1)) Then
MsgBox "CODIGO DE ALUMNO NO EXISTE", 32, "CORREGIR"
GoTo NOV
End If
If ((Text2.Text > 10) Or (Text2.Text < 1)) Then
MsgBox "NUMERO DE OBSERVACION NO EXISTE", 48, "CORREGIR"
GoTo NOV
End If
If Text3.Text = 0 Then
GoTo colmo
End If
If Text3.Text >= FERT Then
MsgBox "OBSERVACION NO EXISTE", 48, "ADVERTENCIA"
GoTo NOV
End If
colmo:
GRABAR_OBSER.MATI12.Row = Text1.Text
GRABAR_OBSER.MATI12.Col = Text2.Text
If GRABAR_OBSER.MATI12.Text = "" Then
MsgBox "NO EXISTE CODIGO PARA LA OBSERVACION No." & Text2.Text, 64, "ADVERTENCIA"
GoTo NOV
End If
GRABAR_OBSER.MATI12.CellFontBold = True
GRABAR_OBSER.MATI12.CellForeColor = RGB(255, 0, 0)
GRABAR_OBSER.MATI12.Text = Text3.Text
Text1.SetFocus
NOV:
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONC25
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONC25:
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONC26
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONC26:
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONC27
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONC27:
End Sub
Private Sub Form_Load()
Text1.MaxLength = 2
Text2.MaxLength = 2
Text3.MaxLength = 2
End Sub
