VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONS_OBSER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE OBSERVACIONES"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9495
   Icon            =   "CONS_OBSER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   8775
      Begin MSFlexGridLib.MSFlexGrid MATI11 
         Height          =   3975
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   7011
         _Version        =   327680
         Rows            =   1
         Cols            =   3
         GridColor       =   12582912
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "REPORTE DE OBSERVACIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   240
      Picture         =   "CONS_OBSER.frx":0442
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "CONS_OBSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MATI11.Row = 0
MATI11.Col = 0
MATI11.ColWidth(0) = 500
MATI11.CellForeColor = RGB(255, 255, 255)
MATI11.CellBackColor = RGB(0, 0, 150)
MATI11.Text = "COD"
MATI11.Col = 1
MATI11.ColWidth(1) = 500
MATI11.CellForeColor = RGB(255, 255, 255)
MATI11.CellBackColor = RGB(0, 0, 150)
MATI11.Text = "IND"
MATI11.Col = 2
MATI11.ColWidth(2) = 7000
MATI11.CellForeColor = RGB(255, 255, 255)
MATI11.CellBackColor = RGB(0, 0, 150)
MATI11.Text = "OBSERVACION"
End Sub
