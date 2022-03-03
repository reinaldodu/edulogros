VERSION 5.00
Begin VB.Form ComentariosDesemp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comentarios por desempeños"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   8
         Top             =   1800
         Width           =   7335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   7335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   7335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SUPERIOR:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ALTO:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BÁSICO:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BAJO:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   540
      End
   End
End
Attribute VB_Name = "ComentariosDesemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NAR = FreeFile
Open Ruta & "comentadesemp.edu" For Output As #NAR
comdpe.bajo = Text1(0).Text
comdpe.basico = Text1(1).Text
comdpe.alto = Text1(2).Text
comdpe.superior = Text1(3).Text
Write #NAR, comdpe.bajo, comdpe.basico, comdpe.alto, comdpe.superior
Close #NAR
Unload Me
End Sub

Private Sub Form_Load()
NAR = FreeFile
Open Ruta & "comentadesemp.edu" For Input As #NAR
Input #NAR, comdpe.bajo, comdpe.basico, comdpe.alto, comdpe.superior
Close #NAR
Text1(0) = comdpe.bajo
Text1(1) = comdpe.basico
Text1(2) = comdpe.alto
Text1(3) = comdpe.superior
End Sub
