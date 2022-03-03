VERSION 5.00
Begin VB.Form ACERCADE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de Edulogros"
   ClientHeight    =   2685
   ClientLeft      =   2550
   ClientTop       =   2385
   ClientWidth     =   4560
   Icon            =   "ACERCA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   390
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "www.educolibre.co"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Bogotá, Colombia. Febrero de 2012"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Creado por:  EducoLibre.co"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1950
   End
   Begin VB.Line Line1 
      DrawMode        =   4  'Mask Not Pen
      X1              =   120
      X2              =   4440
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "Software desarrollado para la sistematización de colegios colombianos - Decreto 1290."
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "EDULOGROS.  V.12.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "ACERCA.frx":27A2
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "ACERCADE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra información acerca de Edulogros."
End Sub
