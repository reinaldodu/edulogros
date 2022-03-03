VERSION 5.00
Begin VB.Form Del_Planeador 
   Caption         =   "Eliminar fecha"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione la fecha a eliminar (dd/mm/aaaa)"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.ListBox List1 
         Height          =   4155
         ItemData        =   "Del_Planeador.frx":0000
         Left            =   240
         List            =   "Del_Planeador.frx":0002
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Del_Planeador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.SelCount = 0 Then
    MsgBox "No ha seleccionado una fecha para eliminar", 16, "Eliminar"
    Exit Sub
End If
MS1 = "Desea eliminar la fecha " & List1.List(List1.ListIndex) & " de la planeación?"
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "Eliminar")
If RESP = vbYes Then
    h = 0
    NAR = FreeFile
    Open Ruta & planeacion_semanal.Label4.Caption & que & lw & ".pln" For Random As #NAR Len = Len(semanal_planeacion)
    While Not EOF(NAR)
        h = h + 1
        Get #NAR, h, semanal_planeacion
        If Trim(semanal_planeacion.fecha) = List1.List(List1.ListIndex) Then
            semanal_planeacion.fecha = ""
            Put #NAR, h, semanal_planeacion
        End If
    Wend
    Close #NAR
    For I = 1 To planeacion_semanal.MTPlan.Rows - 1
        If planeacion_semanal.MTPlan.TextMatrix(I, 0) = List1.List(List1.ListIndex) Then
            planeacion_semanal.MTPlan.RemoveItem I
            Exit For
        End If
    Next I
End If
Unload Me
End Sub
