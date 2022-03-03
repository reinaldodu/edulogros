VERSION 5.00
Begin VB.Form ControlPeriodos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de acceso a  periodos académicos"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6165
   Icon            =   "ControlPeriodos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "DESEMPEÑOS"
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   4935
      Begin VB.CheckBox Check6 
         Caption         =   "FINAL"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         Caption         =   "CUARTO"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.CheckBox Check6 
         Caption         =   "TERCERO"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "SEGUNDO"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox Check6 
         Caption         =   "PRIMERO"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "LOGROS"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      Begin VB.CheckBox Check1 
         Caption         =   "FINAL"
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CUARTO"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TERCERO"
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "SEGUNDO"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "PRIMERO"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "(Seleccione los periodos que desea habilitar)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "ControlPeriodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p1 As String, p2 As String, p3 As String, p4 As String, p5 As String

Private Sub Command1_Click()

'GUARDAR ARCHIVO DE BLOQUEO DE LOGROS
ValoresCK = ""
For J = 0 To 4
    If Check1(J).Value = 1 Then
        ValoresCK = ValoresCK + "1,"
    Else
        ValoresCK = ValoresCK & "0,"
    End If
Next J
VerCHK = Split(ValoresCK, ",")
NAR = FreeFile
Open Ruta & "periodosL.edu" For Output As #NAR
Write #NAR, VerCHK(0), VerCHK(1), VerCHK(2), VerCHK(3), VerCHK(4)
Close #NAR

'GUARDAR ARCHIVO DE BLOQUEO DE DESEMPEÑOS
ValoresCK = ""
For J = 0 To 4
    If Check6(J).Value = 1 Then
        ValoresCK = ValoresCK + "1,"
    Else
        ValoresCK = ValoresCK & "0,"
    End If
Next J
VerCHK = Split(ValoresCK, ",")
NAR = FreeFile
Open Ruta & "periodosD.edu" For Output As #NAR
Write #NAR, VerCHK(0), VerCHK(1), VerCHK(2), VerCHK(3), VerCHK(4)
Close #NAR

Unload Me
End Sub

Private Sub Form_Load()
Dim VerCHK(5) As String

'MOSTRAR INFORMACION DE BLOQUEO DE LOGROS
If Dir(Ruta & "periodosL.edu") <> "" Then
    NAR = FreeFile
    Open Ruta & "periodosL.edu" For Input As #NAR
    Input #NAR, VerCHK(0), VerCHK(1), VerCHK(2), VerCHK(3), VerCHK(4)
    Close #NAR
    For J = 0 To 4
        If VerCHK(J) = "1" Then
            Check1(J).Value = 1
        Else
            Check1(J).Value = 0
        End If
    Next J
End If

'MOSTRAR INFORMACION DE BLOQUEO DE DESEMPEÑOS
If Dir(Ruta & "periodosD.edu") <> "" Then
    NAR = FreeFile
    Open Ruta & "periodosD.edu" For Input As #NAR
    Input #NAR, VerCHK(0), VerCHK(1), VerCHK(2), VerCHK(3), VerCHK(4)
    Close #NAR
    For J = 0 To 4
        If VerCHK(J) = "1" Then
            Check6(J).Value = 1
        Else
            Check6(J).Value = 0
        End If
    Next J
End If

End Sub
