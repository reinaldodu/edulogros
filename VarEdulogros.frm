VERSION 5.00
Begin VB.Form VarEdulogros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Variables Edulogros"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "VarEdulogros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2040
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nombre para periodo:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre para fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre para grupo:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre para estudiante:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre para director:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre para rector:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1410
      End
   End
End
Attribute VB_Name = "VarEdulogros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
vini.VRector = Text1.Text
vini.VDirector = Text2.Text
vini.VEstudiante = Text3.Text
vini.VGrupo = Text4.Text
vini.VFecha = Text5.Text
vini.VPeriodo = Text6.Text
vini.VOp1 = ""
vini.VOp2 = ""
vini.VOp3 = ""
Open Ruta & "VarEdu.edu" For Output As #NAR
Write #NAR, vini.VRector, vini.VDirector, vini.VEstudiante, vini.VGrupo, vini.VFecha, vini.VPeriodo, vini.VOp1, vini.VOp2, vini.VOp3
Close #NAR
Unload Me
End Sub

Private Sub Form_Load()
If Dir(Ruta & "VarEdu.edu") = "" Then
        Exit Sub
End If
Open Ruta & "VarEdu.edu" For Input As #NAR
Input #NAR, vini.VRector, vini.VDirector, vini.VEstudiante, vini.VGrupo, vini.VFecha, vini.VPeriodo, vini.VOp1, vini.VOp2, vini.VOp3
Close #NAR
Text1.Text = vini.VRector
Text2.Text = vini.VDirector
Text3.Text = vini.VEstudiante
Text4.Text = vini.VGrupo
Text5.Text = vini.VFecha
Text6.Text = vini.VPeriodo
End Sub
