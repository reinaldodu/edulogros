VERSION 5.00
Begin VB.Form proyectos 
   Caption         =   "Proyectos"
   ClientHeight    =   9570
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   13020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   8280
         Width           =   11055
      End
      Begin VB.Frame Frame7 
         Caption         =   "Evaluación"
         Height          =   1215
         Left            =   6480
         TabIndex        =   18
         Top             =   6960
         Width           =   6015
         Begin VB.TextBox Text10 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Recursos"
         Height          =   1215
         Left            =   240
         TabIndex        =   17
         Top             =   6960
         Width           =   5895
         Begin VB.TextBox Text9 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   240
            Width           =   5655
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cronograma y actividades"
         Height          =   1215
         Left            =   240
         TabIndex        =   14
         Top             =   5640
         Width           =   12255
         Begin VB.TextBox Text8 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   240
            Width           =   12015
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Metodología"
         Height          =   1215
         Left            =   240
         TabIndex        =   13
         Top             =   4320
         Width           =   12255
         Begin VB.TextBox Text7 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   240
            Width           =   12015
         End
      End
      Begin VB.TextBox Text6 
         Height          =   855
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1680
         Width           =   10575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Competencias"
         Height          =   1575
         Left            =   6360
         TabIndex        =   9
         Top             =   2640
         Width           =   6135
         Begin VB.TextBox Text5 
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ejes temáticos"
         Height          =   1575
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   5895
         Begin VB.TextBox Text4 
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   5655
         End
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1200
         Width           =   10575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   10575
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   10575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   8280
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Objetivos:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Población a trabajar:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Responsables:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del proyecto:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1515
      End
   End
End
Attribute VB_Name = "proyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
