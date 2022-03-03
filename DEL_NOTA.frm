VERSION 5.00
Begin VB.Form DEL_NOTA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   Icon            =   "DEL_NOTA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
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
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   675
      End
   End
End
Attribute VB_Name = "DEL_NOTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "ESCRIBA EL CODIGO A ELIMINAR", 48, "ELIMINAR"
Text1.SetFocus
Exit Sub
End If
If Val(GRABAR_OBSER.Text7.Text) = 1 Then
         MsgBox "IMPOSIBLE BORRAR EL ULTIMO ALUMNO", 32
         Exit Sub
End If
If Val(Text1.Text) > Val(GRABAR_OBSER.Text7.Text) Or (Text1.Text < 1) Then
MsgBox "NO EXISTE CÓDIGO", 32, "ELIMINAR CÓDIGO"
Text1.Text = ""
Text1.SetFocus
Exit Sub
End If
GRABAR_OBSER.MATI12.RemoveItem Val(Text1.Text)
GRABAR_OBSER.Text7.Text = Val(GRABAR_OBSER.Text7.Text) - 1
GRABAR_OBSER.MATI12.Col = 0
For TT = 1 To Val(GRABAR_OBSER.Text7.Text)
GRABAR_OBSER.MATI12.Row = TT
GRABAR_OBSER.MATI12.Text = TT
Next TT
Text1.Text = ""
Text1.SetFocus
VALI4 = False
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Borra un alumno de la lista, escribiendo el código que le corresponde dentro del grupo."
End Sub

Private Sub Form_Load()
Text1.MaxLength = 2
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
