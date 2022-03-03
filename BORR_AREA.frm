VERSION 5.00
Begin VB.Form BORR_AREA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BORRAR AREA."
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "BORR_AREA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   320
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   320
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "GRUPO:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No. AREA:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "BORR_AREA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim argra As areagr
If Text1.Text = "" Then
MsgBox "ESCRIBA EL CODIGO DEL AREA", 32, "BORRAR"
Text1.SetFocus
GoTo GLL58
End If
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NOMBRE DEL GRUPO", 32, "BORRAR"
Text2.SetFocus
GoTo GLL58
End If
Text2.Text = Format(Text2.Text, ">")
RESP = MsgBox("DESEA ELIMINAR ESTA AREA PARA ESTE GRUPO?", vbYesNo + vbQuestion + vbDefaultButton2, "BORRAR AREA")
If RESP = vbYes Then
CLO = 0
NAR = FreeFile
Open "c:\windows\datos\areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
CLO = CLO + 1
Get #NAR, CLO, argra
If ((argra.num_area = Val(Text1.Text)) And (RTrim(argra.nom_grup) = RTrim(Text2.Text))) Then
argra.grado = ""
argra.ih = 0
argra.nom_grup = ""
argra.num_area = 0
argra.num_pro = 0
Put #NAR, CLO, argra
Close #NAR
GoTo osis
End If
Wend
Close #NAR
MsgBox "AREA NO ESTA CREADA PARA ESTE GRUPO", 64, "ADVERTENCIA"
Text1.SetFocus
GoTo GLL58
osis:
MsgBox "EL AREA HA SIDO ELIMINADA", 16, "BORRAR"
End If
GLL58:
End Sub

Private Sub Form_Load()
Text1.MaxLength = 2
Text2.MaxLength = 13
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONC322
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONC322:
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
