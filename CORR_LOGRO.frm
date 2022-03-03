VERSION 5.00
Begin VB.Form CORR_LOGRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CORREGIR LOGRO."
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3015
   Icon            =   "CORR_LOGRO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&GUARDAR"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H0000FFFF&
      Height          =   320
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   2535
      Begin VB.Image Image1 
         Height          =   405
         Left            =   120
         Picture         =   "CORR_LOGRO.frx":0442
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   795
      End
   End
End
Attribute VB_Name = "CORR_LOGRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim logru As logris
If Text1.Text = "" Then
MsgBox "ESCRIBA EL CODIGO DE LA OBSERVACION", 32, "CORREGIR"
Text1.SetFocus
GoTo segg
End If
LOGROS.Show
If LOGROS.Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DEL AREA Y PRESIONE OK", 48, "CORREGIR"
LOGROS.Text1.SetFocus
GoTo segg
End If
If (Text1.Text > LOGROS.Text5.Text) Or (Text1.Text < 1) Then
MsgBox "NO EXISTE CODIGO DE OBSERVACION", 32, "CORREGIR"
Text1.SetFocus
GoTo segg
End If
NAR = FreeFile
Open "c:\windows\datos\" & fl & ser & Val(LOGROS.Text1.Text) & lw & ".lgr" For Random As #NAR Len = Len(logru)
Get #NAR, Text1.Text, logru
Close #NAR
LOGROS.Show
LOGROS.Text4.Text = RTrim(logru.observ)
LOGROS.Text3.Text = Text1.Text
LOGROS.Combo3.Text = logru.indicador
LOGROS.Command1.Enabled = False
LOGROS.Command3.Enabled = False
LOGROS.Command4.Enabled = False
LOGROS.Text4.SetFocus
ABC = True
segg:
End Sub

Private Sub Command2_Click()
Dim logru As logris
If ABC = False Then
MsgBox "PRESIONE PRIMERO OK", 64, "CORREGIR"
GoTo LULU
End If
LOGROS.Show
If LOGROS.Text3.Text = "" Then
Unload LOGROS
MsgBox "NO HAY INFORMACION PARA GUARDAR", 16, "CORREGIR"
GoTo LULU
End If
RESP = MsgBox("DESEA GUARDAR LA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
NAR = FreeFile
Open "c:\windows\datos\" & fl & ser & Val(LOGROS.Text1.Text) & lw & ".lgr" For Random As #NAR Len = Len(logru)
logru.indicador = Format(LOGROS.Combo3.Text, ">")
logru.observ = LOGROS.Text4.Text
Put #NAR, LOGROS.Text3.Text, logru
Close #NAR
End If
ABC = False
LULU:
Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONC23
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONC23:
End Sub
Private Sub Form_Load()
Text1.MaxLength = 2
ABC = False
End Sub
