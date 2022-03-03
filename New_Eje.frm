VERSION 5.00
Begin VB.Form New_Eje 
   Caption         =   "Agregar ejes temáticos y contenidos"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.Frame Frame2 
         Caption         =   "Contenidos (de doble clic para modificar el contenido seleccionado)"
         Height          =   4215
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   9615
         Begin VB.CommandButton Command4 
            Caption         =   "Pegar"
            Height          =   375
            Left            =   8040
            TabIndex        =   8
            Top             =   3720
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Copiar"
            Height          =   375
            Left            =   6120
            TabIndex        =   7
            Top             =   3720
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Eliminar contenido"
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   3720
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar contenido"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   3720
            Width           =   1815
         End
         Begin VB.ListBox List_Contenido 
            Height          =   3375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   9375
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Eje temático:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1110
      End
   End
End
Attribute VB_Name = "New_Eje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Mod_Vr = InputBox("Escriba el contenido que desea agregar", "Agregar contenido")
If Trim(Mod_Vr) = "" Then
    MsgBox "No escribió el contenido", 64, "Agregar contenido"
    Exit Sub
End If
List_Contenido.AddItem Trim(Mod_Vr)
'Frame2.Caption = List_Contenido.ListCount
End Sub

Private Sub Command2_Click()
'For I = 0 To List_Contenido.ListCount - 1
If List_Contenido.SelCount Then
    List_Contenido.RemoveItem (List_Contenido.ListIndex)
End If
'Next I
End Sub

Private Sub Command3_Click()
If List_Contenido.ListCount = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA COPIAR", 16, "COPIAR"
    Exit Sub
End If
Screen.MousePointer = 11
Clipboard.Clear
cop = ""
For X = 0 To List_Contenido.ListCount - 1
    INDI = List_Contenido.List(X)
    cop = cop + INDI & vbCrLf
Next X
Close #NAR
Clipboard.SetText cop
Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()

ContList = Clipboard.GetText
ArrCont = Split(ContList, vbCrLf)
For J = 0 To UBound(ArrCont)
List_Contenido.AddItem ArrCont(J)
Next

End Sub

Private Sub Command5_Click()
If Trim(Text1.Text) = "" Then
    MsgBox "No ha escrito el eje temático", 16, "ADVERTENCIA"
    Exit Sub
End If
If List_Contenido.ListCount = 0 Then
    MsgBox "Debe agregar contenidos al eje temático", 16, "ADVERTENCIA"
    Exit Sub
End If
If ValiModifica = True Then
    TT = Val(TTT)
    Ejes_Contenidos.MTEjes.TextMatrix(TT, 1) = ""
    Ejes_Contenidos.MTEjes.TextMatrix(TT, 2) = ""
    Ejes_Contenidos.MTEjes.RowHeight(TT) = 240
Else
    TT = Ejes_Contenidos.MTEjes.Rows
    Ejes_Contenidos.MTEjes.Rows = TT + 1
    Ejes_Contenidos.MTEjes.TextMatrix(TT, 0) = TT
End If
Ejes_Contenidos.MTEjes.Row = TT
Ejes_Contenidos.MTEjes.Col = 1
Ejes_Contenidos.MTEjes.CellFontBold = True
Ejes_Contenidos.MTEjes.Text = Trim(Text1.Text)

For J = 0 To List_Contenido.ListCount - 1
    Ejes_Contenidos.MTEjes.TextMatrix(TT, 2) = Ejes_Contenidos.MTEjes.TextMatrix(TT, 2) & List_Contenido.List(J) & vbCrLf
    Ejes_Contenidos.MTEjes.RowHeight(TT) = Ejes_Contenidos.MTEjes.RowHeight(TT) + 240
Next J
VALI380 = False
Unload Me
End Sub

Private Sub Form_Load()
If ValiModifica = True Then
    Text1.Text = Trim(Ejes_Contenidos.MTEjes.TextMatrix(TTT, 1))
    ArrCont = Split(Trim(Ejes_Contenidos.MTEjes.TextMatrix(TTT, 2)), vbCrLf)
    For r = 0 To UBound(ArrCont)
        If ArrCont(r) <> "" Then
            List_Contenido.AddItem Trim(ArrCont(r))
        End If
    Next r
End If
End Sub

Private Sub List_Contenido_DblClick()
Mod_Vr = InputBox("", "Modificar contenido", List_Contenido.List(List_Contenido.ListIndex))
If Trim(Mod_Vr) <> "" Then
    List_Contenido.List(List_Contenido.ListIndex) = Mod_Vr
End If
End Sub

