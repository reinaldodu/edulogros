VERSION 5.00
Begin VB.Form New_Competencia 
   Caption         =   "Agregar competencia"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.Frame Frame2 
         Caption         =   "Logros (Oprima <Ctrl> para seleccionar varios logros o <Shift> para seleccionar en bloque)"
         Height          =   2775
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   9855
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   6
            Top             =   240
            Width           =   9615
         End
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1200
         MaxLength       =   700
         TabIndex        =   4
         Top             =   840
         Width           =   8895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Competencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   540
      End
   End
End
Attribute VB_Name = "New_Competencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim kk As String
If Trim(Text1.Text) = "" Then
    MsgBox "No ha escrito el código de la comptencia", 16, "ADVERTENCIA"
    Exit Sub
End If
If Trim(Text2.Text) = "" Then
    MsgBox "No ha escrito la comptencia", 16, "ADVERTENCIA"
    Exit Sub
End If
If ValiModifica = False Then
    TT = Competencias.MTComp.Rows
    Competencias.MTComp.Rows = TT + 1
    Competencias.MTComp.TextMatrix(TT, 0) = TT
    Competencias.MTComp.TextMatrix(TT, 1) = Trim(Text1.Text)
    Competencias.MTComp.TextMatrix(TT, 2) = Trim(Text2.Text)
Else
    'Competencias.MTComp.TextMatrix(TTT, 0) = TTT
    Competencias.MTComp.TextMatrix(TTT, 1) = Trim(Text1.Text)
    Competencias.MTComp.TextMatrix(TTT, 2) = Trim(Text2.Text)
End If

Y = 0
For I = 0 To List1.ListCount - 1
    If List1.Selected(I) = True Then
        kk = kk & (I + 1) & ","
        Y = Y + 1
    End If
Next I
If Y = 0 Then
    MsgBox "Debe seleccionar por lo menos un logro", 16, "ADVERTENCIA"
    Exit Sub
End If
If ValiModifica = False Then
    Competencias.MTComp.TextMatrix(TT, 3) = kk
Else
    Competencias.MTComp.TextMatrix(TTT, 3) = kk
End If
VALI180 = False
Unload Me
End Sub

Private Sub Form_Load()
If ValiModifica = True Then
    Text1.Text = Competencias.MTComp.TextMatrix(TTT, 1)
    Text2.Text = Competencias.MTComp.TextMatrix(TTT, 2)
    ArrLogros = Split(Competencias.MTComp.TextMatrix(TTT, 3), ",")
End If
List1.Clear
w = 0
t = 0
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    w = w + 1
    Get #NAR, w, logru
    If Trim(logru.indicador) = "L" Then
        List1.AddItem Trim(logru.observ)
        If ValiModifica = True Then
            For h = 0 To UBound(ArrLogros)
                If t + 1 = Val(ArrLogros(h)) Then
                    List1.Selected(t) = True
                End If
            Next h
        End If
        t = t + 1
    End If
Wend
Close #NAR
End Sub
