VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Control_Proyectos 
   Caption         =   "Proyectos por áreas y proyectos transversales"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proyectos transversales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   10215
      Begin VB.CommandButton Command7 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   2280
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   2280
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1500
      End
      Begin MSFlexGridLib.MSFlexGrid MT_PYtransversal 
         Height          =   1935
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proyectos por áreas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.CommandButton Command6 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2280
         Width           =   1500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   2280
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1500
      End
      Begin MSFlexGridLib.MSFlexGrid MT_PYareas 
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
      End
   End
End
Attribute VB_Name = "Control_Proyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NewPyArea = True
New_Proyecto.Show 1
End Sub

Private Sub Command2_Click()
If Val(MT_PYareas.Rows - 1) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA ELIMINAR", 64
    Exit Sub
End If
TTT = InputBox("Escriba el número del proyecto que desea eliminar", "Eliminar")
If TTT = "" Then
    MsgBox "No escribió el número del proyecto", 64, "Eliminar"
    Exit Sub
End If
If Val(TTT) > Val(MT_PYareas.Rows - 1) Or (Val(TTT) < 1) Then
    MsgBox "No existe este número de proyecto", 32, "Eliminar"
    Exit Sub
End If
MT_PYareas.RemoveItem Val(TTT)
For I = 1 To Val(MT_PYareas.Rows - 1)
    MT_PYareas.TextMatrix(I, 0) = I
Next I
End Sub

Private Sub Command3_Click()
NewPyArea = False
New_Proyecto.Show 1
End Sub

Private Sub Command4_Click()
If Val(MT_PYtransversal.Rows - 1) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA ELIMINAR", 64
    Exit Sub
End If
TTT = InputBox("Escriba el número del proyecto que desea eliminar", "Eliminar")
If TTT = "" Then
    MsgBox "No escribió el número del proyecto", 64, "Eliminar"
    Exit Sub
End If
If Val(TTT) > Val(MT_PYtransversal.Rows - 1) Or (Val(TTT) < 1) Then
    MsgBox "No existe este número de proyecto", 32, "Eliminar"
    Exit Sub
End If
MT_PYtransversal.RemoveItem Val(TTT)
For I = 1 To Val(MT_PYtransversal.Rows - 1)
    MT_PYtransversal.TextMatrix(I, 0) = I
Next I
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
If MT_PYareas.Rows <> 1 Then
    If Dir(Ruta & "proyectos1.pyt") <> "" Then
       Kill Ruta & "proyectos1.pyt"
    End If
    NAR = FreeFile
    For I = 1 To MT_PYareas.Rows - 1
        aliasg = MT_PYareas.TextMatrix(I, 1) & "&&" & MT_PYareas.TextMatrix(I, 2)
        Open Ruta & "proyectos1.pyt" For Append As #NAR
        Write #NAR, aliasg
        Close #NAR
    Next I
Else
    MsgBox "No existe información para guardar", 16, "Proyectos por áreas"
End If
End Sub

Private Sub Command7_Click()
If MT_PYtransversal.Rows <> 1 Then
    If Dir(Ruta & "proyectos2.pyt") <> "" Then
       Kill Ruta & "proyectos2.pyt"
    End If
    NAR = FreeFile
    For I = 1 To MT_PYtransversal.Rows - 1
        aliasg = MT_PYtransversal.TextMatrix(I, 1) & "&&" & MT_PYtransversal.TextMatrix(I, 2)
        Open Ruta & "proyectos2.pyt" For Append As #NAR
        Write #NAR, aliasg
        Close #NAR
    Next I
Else
    MsgBox "No existe información para guardar", 16, "Proyectos transversales"
End If
End Sub

Private Sub Form_Load()
MT_PYareas.Row = 0
MT_PYareas.Col = 0
MT_PYareas.ColWidth(0) = 400
MT_PYareas.CellFontBold = True
MT_PYareas.CellForeColor = RGB(255, 255, 255)
MT_PYareas.CellBackColor = RGB(0, 0, 150)
MT_PYareas.Text = "No."
MT_PYareas.Col = 1
MT_PYareas.ColWidth(1) = 6000
MT_PYareas.CellFontBold = True
MT_PYareas.CellForeColor = RGB(255, 255, 255)
MT_PYareas.CellBackColor = RGB(0, 0, 150)
MT_PYareas.Text = "NOMBRE DEL PROYECTO"
MT_PYareas.Col = 2
MT_PYareas.ColWidth(2) = 3000
MT_PYareas.CellFontBold = True
MT_PYareas.CellForeColor = RGB(255, 255, 255)
MT_PYareas.CellBackColor = RGB(0, 0, 150)
MT_PYareas.Text = "PROFESOR RESPONSABLE"


MT_PYtransversal.Row = 0
MT_PYtransversal.Col = 0
MT_PYtransversal.ColWidth(0) = 400
MT_PYtransversal.CellFontBold = True
MT_PYtransversal.CellForeColor = RGB(255, 255, 255)
MT_PYtransversal.CellBackColor = RGB(0, 0, 150)
MT_PYtransversal.Text = "No."
MT_PYtransversal.Col = 1
MT_PYtransversal.ColWidth(1) = 6000
MT_PYtransversal.CellFontBold = True
MT_PYtransversal.CellForeColor = RGB(255, 255, 255)
MT_PYtransversal.CellBackColor = RGB(0, 0, 150)
MT_PYtransversal.Text = "NOMBRE DEL PROYECTO"
MT_PYtransversal.Col = 2
MT_PYtransversal.ColWidth(2) = 3000
MT_PYtransversal.CellFontBold = True
MT_PYtransversal.CellForeColor = RGB(255, 255, 255)
MT_PYtransversal.CellBackColor = RGB(0, 0, 150)
MT_PYtransversal.Text = "PROFESOR RESPONSABLE"

End Sub
