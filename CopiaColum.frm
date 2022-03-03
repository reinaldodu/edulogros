VERSION 5.00
Begin VB.Form CopiaColum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2310
   Icon            =   "CopiaColum.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copiar columna"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton ColumCop 
         Caption         =   "Teléfono"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1815
      End
      Begin VB.OptionButton ColumCop 
         Caption         =   "Acudiente"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton ColumCop 
         Caption         =   "Edad"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton ColumCop 
         Caption         =   "Fecha de nacimiento"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton ColumCop 
         Caption         =   "Apellidos y Nombres"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton ColumCop 
         Caption         =   "Carnet"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Elijenum 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "CopiaColum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ColumCop_Click(Index As Integer)
Elijenum.Caption = Index
End Sub

Private Sub Command1_Click()
Clipboard.Clear
cop = ""
y = Elijenum.Caption
If y <> 1 Then
    For X = 1 To (ARBOL.MATI9.Rows - 1)
        If y = 0 Then
            ape = Trim(ARBOL.MATI9.TextMatrix(X, y + 1))
        Else
            ape = Trim(ARBOL.MATI9.TextMatrix(X, y + 2))
        End If
        cop = cop + Trim(ape) & vbCrLf
    Next X
Else
    For X = 1 To (ARBOL.MATI9.Rows - 1)
        ape = Trim(ARBOL.MATI9.TextMatrix(X, y + 1))
        nom = Trim(ARBOL.MATI9.TextMatrix(X, y + 2))
        cop = cop + Trim(ape & " " & nom) & vbCrLf
    Next X
End If
Clipboard.SetText cop
Unload Me
End Sub
