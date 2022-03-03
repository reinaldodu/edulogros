VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PASSWS_PROFES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Passwords de profesores"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "PASSWS_PROFES.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin MSFlexGridLib.MSFlexGrid MATI14 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   1
      End
   End
End
Attribute VB_Name = "PASSWS_PROFES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
