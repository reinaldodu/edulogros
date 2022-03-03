VERSION 5.00
Begin VB.Form Est_Grado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.ListBox List1 
         Height          =   5325
         Left            =   240
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Est_Grado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CONT = 1
CURSOS.MATI3.Rows = 1
For I = 0 To List1.ListCount - 1
    If List1.Selected(I) = True Then
        CURSOS.MATI3.Rows = CONT + 1
        CURSOS.MATI3.TextMatrix(CONT, 0) = CONT
        CURSOS.MATI3.TextMatrix(CONT, 1) = Val(Right(List1.List(I), 4))
        CURSOS.MATI3.TextMatrix(CONT, 2) = Left(List1.List(I), Len(List1.List(I)) - 6)
        CONT = CONT + 1
    End If
Next I
CURSOS.Text1.Text = CONT - 1
Unload Me
End Sub
