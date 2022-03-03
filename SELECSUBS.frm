VERSION 5.00
Begin VB.Form SELECSUBS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación y consulta de Subsistemas"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "SELECSUBS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Verificar"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      ToolTipText     =   "Verificar las fechas de actualización y baja disco boletines de los subsistemas"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Co&nsultar"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Consultar los grupos que le pertenecen al Subsistema seleccionado"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Crear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Agregar o eliminar  grupo(s) para el Subsistema seleccionado"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el Subsistema"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.ListBox List1 
         Height          =   2010
         ItemData        =   "SELECSUBS.frx":0442
         Left            =   120
         List            =   "SELECSUBS.frx":0464
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "SELECSUBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.SelCount = 0 Then
    MsgBox "Seleccione primero un Subsistema"
    Exit Sub
End If
AREASTEC.Caption = "Grupos - " & Format(List1.List(List1.ListIndex), "<")
Unload Me
AREASTEC.Show 1
End Sub

Private Sub Command2_Click()
If List1.SelCount = 0 Then
    MsgBox "Seleccione primero un Subsistema"
    Exit Sub
End If
If Dir(Ruta & "subsis" & (List1.ListIndex) + 1 & ".sub") = "" Then
    MsgBox "No existen grupos creados para este Subsistema", 64, "Consultar"
    Exit Sub
End If
CONSARESUB.Frame1.Caption = "Grupos Subsistema No." & List1.ListIndex + 1
CONSARESUB.Show 1
End Sub

Private Sub Command3_Click()
If Dir(Ruta & "infosub.edu") = "" Then
    MsgBox "No hay información disponible", 64, "Verificar"
    Exit Sub
End If
VERISUBSIST.Show 1
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Creación, consulta y verificación de Subsistemas."
End Sub
