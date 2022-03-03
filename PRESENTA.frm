VERSION 5.00
Begin VB.Form PRESENTA 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3120
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PRESENTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5880
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   2280
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Plataforma Windows 95 / 98 / 2000 / NT / XP / Vista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         TabIndex        =   3
         Top             =   960
         Width           =   2610
      End
      Begin VB.Image imgLogo 
         Height          =   2745
         Left            =   0
         Picture         =   "PRESENTA.frx":000C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Versión  9.11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   2
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Software desarrollado para la sistematización de colegios."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "PRESENTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Ruta = "c:\windows\datos\"
'Lee el archivo BD.txt que contiene la ruta de los datos y se guarda en la variable Ruta.
If Dir(App.Path & "\BD.txt") = "" Then
    MsgBox "NO EXISTE EL ARCHIVO BD.txt", 48
    End
Else
    NAR = FreeFile
    Open (App.Path & "\BD.txt") For Input As #NAR
    Input #NAR, Ruta
    Close #NAR
End If
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
Timer1.Interval = 1800
End Sub

Private Sub Timer1_Timer()
Unload Me
If (Dir(Ruta & "CONMATRI.EDU") = "") Or (Dir(Ruta & "CONT.EDU") = "") Or (Dir(Ruta & "CONTPRO.EDU") = "") Or (Dir(Ruta & "INICIAL.EDU") = "") Or (Dir(Ruta & "LEYENDA.EDU") = "") Then
    MsgBox "NO SE ENCUENTRAN COMPLETOS LOS ARCHIVOS DE INICIO DEL PROGRAMA, POR FAVOR COMUNIQUESE CON SU PROVEEDOR.", 16
    End
End If
CONTRASEÑA.Show 1
Timer1.Enabled = False
End Sub
