VERSION 5.00
Begin VB.Form Est_Grado2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.ListBox List1 
         Height          =   5325
         Left            =   240
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Est_Grado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'AGREGAR ESTUDIANTES AL GRUPO
CONT = CONS_GRUP.MATI9.Rows
For I = 0 To List1.ListCount - 1
    If List1.Selected(I) = True Then
        CONS_GRUP.MATI9.Rows = CONT + 1
        'CONS_GRUP.MATI9.TextMatrix(CONT, 0) = CONT
        CONS_GRUP.MATI9.TextMatrix(CONT, 1) = Val(Right(List1.List(I), 4))
        CONS_GRUP.MATI9.TextMatrix(CONT, 2) = Left(List1.List(I), Len(List1.List(I)) - 6)
        CONT = CONT + 1
    End If
Next I

' ORDENAR ALFABETICAMENTE
CONS_GRUP.MATI9.Col = 2
CONS_GRUP.MATI9.Sort = 5
For I = 1 To CONS_GRUP.MATI9.Rows - 1
    CONS_GRUP.MATI9.TextMatrix(I, 0) = I
Next I

' GUARDAR LISTADO DE ESTUDIANTES
Kill Ruta & RESC & ".gru"
For we = 1 To CONS_GRUP.MATI9.Rows - 1
    Open Ruta & RESC & ".gru" For Random As #NAR Len = Len(alugru)
    alugru.num_carnet = CONS_GRUP.MATI9.TextMatrix(we, 1)
    Put #NAR, we, alugru
    Close #NAR
    
    aluper.grupo = RTrim(RESC)
    Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
    Put #NAR, Val(CONS_GRUP.MATI9.TextMatrix(we, 1)), aluper
    Close #NAR
Next we

CONS_GRUP.Text4.Text = CONT - 1
Unload Me
End Sub
