VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VERI_PROSIS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación de disco-datos exitosos"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "VERI_PROSIS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid MATI28 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   4
      End
   End
End
Attribute VB_Name = "VERI_PROSIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra para que periodo y fecha fueron bajados al sistema principal los disco-datos de profesores."
End Sub
