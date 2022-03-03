VERSION 5.00
Begin VB.Form CORR_AREGRA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CORREGIR AREAS-GRADO"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4215
   Icon            =   "CORR_AREGRA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   320
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&GUARDAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   320
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   3735
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO DEL AREA:"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRE DEL GRUPO:"
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
         TabIndex        =   5
         Top             =   840
         Width           =   2010
      End
   End
End
Attribute VB_Name = "CORR_AREGRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
