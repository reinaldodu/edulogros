VERSION 5.00
Begin VB.Form Cambiar_Base 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Base de Datos"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   Icon            =   "Cambiar_Base.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.ComboBox cambia 
         Height          =   315
         ItemData        =   "Cambiar_Base.frx":0442
         Left            =   720
         List            =   "Cambiar_Base.frx":04C4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   330
      End
   End
End
Attribute VB_Name = "Cambiar_Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If cambia.Text <> "Actual" Then
    If Dir(Ruta & "historia\" & cambia.Text & "\", vbDirectory) <> "" Then
'       Ruta = "c:\windows\datos\historia\" & cambia.Text & "\"
        Ruta = Ruta & "historia\" & cambia.Text & "\"
'        ENTRADA.coprofeco.Enabled = False
'        ENTRADA.cotecaredis.Enabled = False
        ENTRADA.histo.Enabled = False
        ENTRADA.inian.Enabled = False
    Else
        MsgBox "No existe historial para este año", 16, "Cambiar"
        Exit Sub
    End If
Else
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
'    ENTRADA.coprofeco.Enabled = True
'    ENTRADA.cotecaredis.Enabled = True
    ENTRADA.histo.Enabled = True
    ENTRADA.inian.Enabled = True
End If
MsgBox "La Base de Datos ha cambiado", 64, "Cambiar"
ENTRADA.Caption = "EDULOGROS - (" & cambia.Text & ")"
Unload Me
End Sub

Private Sub Form_Load()
cambia.Text = cambia.List(0)
End Sub
