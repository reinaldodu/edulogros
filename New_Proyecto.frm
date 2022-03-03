VERSION 5.00
Begin VB.Form New_Proyecto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo proyecto"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el profesor responsable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6015
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre del proyecto:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1830
   End
End
Attribute VB_Name = "New_Proyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Trim(Text1.Text) = "" Then
    MsgBox "No ha escrito el nombre del proyecto", 16, "Nuevo proyecto"
    Exit Sub
End If
If List1.SelCount = 0 Then
    MsgBox "Debe seleccionar un profesor", 16, "Nuevo proyecto"
    Exit Sub
End If

If NewPyArea = True Then
    Control_Proyectos.MT_PYareas.Rows = Control_Proyectos.MT_PYareas.Rows + 1
    Control_Proyectos.MT_PYareas.TextMatrix(Control_Proyectos.MT_PYareas.Rows - 1, 0) = Control_Proyectos.MT_PYareas.Rows - 1
    Control_Proyectos.MT_PYareas.TextMatrix(Control_Proyectos.MT_PYareas.Rows - 1, 1) = Text1
    Control_Proyectos.MT_PYareas.TextMatrix(Control_Proyectos.MT_PYareas.Rows - 1, 2) = List1.List(List1.ListIndex)
Else
    Control_Proyectos.MT_PYtransversal.Rows = Control_Proyectos.MT_PYtransversal.Rows + 1
    Control_Proyectos.MT_PYtransversal.TextMatrix(Control_Proyectos.MT_PYtransversal.Rows - 1, 0) = Control_Proyectos.MT_PYtransversal.Rows - 1
    Control_Proyectos.MT_PYtransversal.TextMatrix(Control_Proyectos.MT_PYtransversal.Rows - 1, 1) = Text1
    Control_Proyectos.MT_PYtransversal.TextMatrix(Control_Proyectos.MT_PYtransversal.Rows - 1, 2) = List1.List(List1.ListIndex)
End If
Unload Me
End Sub

Private Sub Form_Load()
r = 0
NAR = FreeFile
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
While Not EOF(NAR)
    r = r + 1
    Get #NAR, r, profe
Wend
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
For I = 1 To (r - 1)
    Get #NAR, I, profe
    If Trim(profe.nombres) <> "" And Trim(profe.apellidos) <> "" Then
        List1.AddItem Trim(profe.nombres) & " " & Trim(profe.apellidos) & " - " & Format(I, "000")
    End If
Next I
Close #NAR
End Sub
