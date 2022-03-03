VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONSVARIAS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas opcionales"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "CONSVARIAS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   4320
      Picture         =   "CONSVARIAS.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Copia la lista de alumnos encontrados"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   3360
      Picture         =   "CONSVARIAS.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ordena los registros alfabeticamente por apellidos"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   5280
      Picture         =   "CONSVARIAS.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprime la información de la lista"
      Top             =   4080
      Width           =   735
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00800000&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid MATRICON 
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
      End
   End
End
Attribute VB_Name = "CONSVARIAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
IMPR_CVARI.Show 1
End Sub

Private Sub Command2_Click()
MATRICON.Col = 1
MATRICON.Sort = 5
End Sub

Private Sub Command3_Click()
Clipboard.Clear
cop = ""
For X = 1 To (MATRICON.Rows - 1)
        ape = RTrim(MATRICON.TextMatrix(X, 0))
        nom = RTrim(MATRICON.TextMatrix(X, 1))
        If X < 10 Then
           cop = cop + LTrim(Str(X) & "   - " & ape & "  " & nom) & vbCrLf
        Else
           cop = cop + LTrim(Str(X) & " - " & ape & "  " & nom) & vbCrLf
        End If
Next X
Clipboard.SetText cop
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Resultados de la " & Format(Frame1.Caption, "<")
End Sub
