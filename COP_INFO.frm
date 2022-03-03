VERSION 5.00
Begin VB.Form COP_INFO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar información"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "COP_INFO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Copiar"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Imprime los campos seleccionados"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2760
      ItemData        =   "COP_INFO.frx":0442
      Left            =   240
      List            =   "COP_INFO.frx":046A
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Campos que se imprimen"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "COP_INFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim ini As inicio
Clipboard.Clear
cop = ""
For X = 0 To (ARBOL.MATI9.Rows - 1)
    'ape = Trim(ARBOL.MATI9.TextMatrix(X, Y + 1))
    'nom = Trim(ARBOL.MATI9.TextMatrix(X, Y + 2))
    'cop = cop + Trim(ape & " " & nom) & vbTab
    'For I = 0 To 9
    For I = 0 To 11
        If List1.Selected(I) = True Then
            'cop = cop + Trim(ARBOL.MATI9.TextMatrix(X, Y + I + 3)) & vbTab
            cop = cop + Trim(ARBOL.MATI9.TextMatrix(X, Y + I + 1)) & vbTab
        End If
    Next I
    cop = cop + vbCrLf
Next X
Clipboard.SetText cop
Unload Me
End Sub
